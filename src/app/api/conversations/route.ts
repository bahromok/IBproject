import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/conversations - List conversations for current user
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const conversations = await db.conversation.findMany({
      where: {
        OR: [{ participant1Id: user.id }, { participant2Id: user.id }],
      },
      orderBy: { lastMessageAt: 'desc' },
      include: {
        participant1: {
          select: {
            id: true,
            clientProfile: { select: { displayName: true } },
            freelancerProfile: { select: { displayName: true, avatar: true } },
          },
        },
        participant2: {
          select: {
            id: true,
            clientProfile: { select: { displayName: true } },
            freelancerProfile: { select: { displayName: true, avatar: true } },
          },
        },
        contract: {
          select: { id: true, title: true },
        },
        _count: {
          select: {
            messages: {
              where: {
                isRead: false,
                senderId: { not: user.id },
              },
            },
          },
        },
      },
    })

    // Identity protection: return the OTHER user's displayName
    const safeConversations = conversations.map((conv) => {
      const isParticipant1 = conv.participant1Id === user.id
      const otherUser = isParticipant1 ? conv.participant2 : conv.participant1
      const displayName =
        otherUser.clientProfile?.displayName ||
        otherUser.freelancerProfile?.displayName ||
        'Anonymous'
      const avatar =
        otherUser.freelancerProfile?.avatar ||
        null

      return {
        id: conv.id,
        contract: conv.contract,
        otherUser: { id: otherUser.id, displayName, avatar },
        unreadCount: conv._count.messages,
        lastMessageAt: conv.lastMessageAt,
        createdAt: conv.createdAt,
      }
    })

    return apiSuccess({ conversations: safeConversations })
  } catch (error: any) {
    console.error('Conversations list error:', error)
    return apiError('Failed to fetch conversations', 500)
  }
}

// POST /api/conversations - Create new conversation
export async function POST(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const body = await request.json()
    const { user2Id, jobId } = body

    if (!user2Id) return apiError('User ID is required')
    if (user2Id === user.id) return apiError('Cannot create a conversation with yourself')

    // Check if user2 exists
    const otherUser = await db.user.findUnique({ where: { id: user2Id } })
    if (!otherUser) return apiError('User not found', 404)

    // Check for existing conversation
    const existing = await db.conversation.findFirst({
      where: {
        OR: [
          { participant1Id: user.id, participant2Id: user2Id },
          { participant1Id: user2Id, participant2Id: user.id },
        ],
      },
    })

    if (existing) {
      return apiSuccess({ conversation: existing, message: 'Conversation already exists' })
    }

    const conversation = await db.conversation.create({
      data: {
        participant1Id: user.id,
        participant2Id: user2Id,
        contractId: jobId || null,
      },
    })

    return apiSuccess({ conversation }, 201)
  } catch (error: any) {
    console.error('Conversation creation error:', error)
    return apiError('Failed to create conversation', 500)
  }
}
