import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyNewMessage } from '@/lib/notify'

// GET /api/conversations/[id]/messages - Get messages for a conversation
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: conversationId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const conversation = await db.conversation.findUnique({
      where: { id: conversationId },
    })

    if (!conversation) return apiError('Conversation not found', 404)
    if (conversation.participant1Id !== user.id && conversation.participant2Id !== user.id) {
      return apiError('You do not have access to this conversation', 403)
    }

    const { searchParams } = new URL(request.url)
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '50')

    const [messages, total] = await Promise.all([
      db.message.findMany({
        where: { conversationId },
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'asc' },
        select: {
          id: true,
          content: true,
          messageType: true,
          fileUrl: true,
          fileName: true,
          isRead: true,
          createdAt: true,
          senderId: true,
          sender: {
            select: {
              id: true,
              clientProfile: { select: { displayName: true } },
              freelancerProfile: { select: { displayName: true, avatar: true } },
            },
          },
        },
      }),
      db.message.count({ where: { conversationId } }),
    ])

    // Identity protection: only show displayName
    const safeMessages = messages.map((msg) => ({
      id: msg.id,
      content: msg.content,
      messageType: msg.messageType,
      fileUrl: msg.fileUrl,
      fileName: msg.fileName,
      isRead: msg.isRead,
      createdAt: msg.createdAt,
      senderId: msg.senderId,
      isOwnMessage: msg.senderId === user.id,
      sender: {
        id: msg.sender.id,
        displayName:
          msg.sender.clientProfile?.displayName ||
          msg.sender.freelancerProfile?.displayName ||
          'Anonymous',
        avatar:
          msg.sender.freelancerProfile?.avatar ||
          null,
      },
    }))

    // Mark messages as read
    await db.message.updateMany({
      where: {
        conversationId,
        senderId: { not: user.id },
        isRead: false,
      },
      data: { isRead: true },
    })

    return apiSuccess({
      messages: safeMessages,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Messages fetch error:', error)
    return apiError('Failed to fetch messages', 500)
  }
}

// POST /api/conversations/[id]/messages - Send message
export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: conversationId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const conversation = await db.conversation.findUnique({
      where: { id: conversationId },
    })

    if (!conversation) return apiError('Conversation not found', 404)
    if (conversation.participant1Id !== user.id && conversation.participant2Id !== user.id) {
      return apiError('You do not have access to this conversation', 403)
    }

    const body = await request.json()
    const { content, messageType, fileUrl, fileName } = body

    if (!content || !content.trim()) return apiError('Message content is required')

    const message = await db.message.create({
      data: {
        conversationId,
        senderId: user.id,
        content: content.trim(),
        messageType: messageType || 'TEXT',
        ...(fileUrl && { fileUrl }),
        ...(fileName && { fileName }),
      },
    })

    // Update last message timestamp
    await db.conversation.update({
      where: { id: conversationId },
      data: { lastMessageAt: new Date() },
    })

    // Notify the other user (fire-and-forget)
    const recipientId = conversation.participant1Id === user.id ? conversation.participant2Id : conversation.participant1Id
    const senderProfile = user.role === 'FREELANCER'
      ? await db.freelancerProfile.findUnique({ where: { userId: user.id }, select: { displayName: true } })
      : await db.clientProfile.findUnique({ where: { userId: user.id }, select: { displayName: true } })
    const senderName = senderProfile?.displayName || 'Someone'
    notifyNewMessage(recipientId, conversationId, senderName)

    return apiSuccess({ message }, 201)
  } catch (error: any) {
    console.error('Message send error:', error)
    return apiError('Failed to send message', 500)
  }
}
