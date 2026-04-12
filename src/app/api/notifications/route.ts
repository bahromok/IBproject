import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/notifications - Get notifications for current user
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '30')
    const unreadOnly = searchParams.get('unreadOnly') === 'true'

    const where: any = { userId: user.id }
    if (unreadOnly) where.isRead = false

    const [notifications, total, unreadCount] = await Promise.all([
      db.notification.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
      }),
      db.notification.count({ where }),
      db.notification.count({ where: { userId: user.id, isRead: false } }),
    ])

    return apiSuccess({
      notifications,
      unreadCount,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Notifications fetch error:', error)
    return apiError('Failed to fetch notifications', 500)
  }
}

// PATCH /api/notifications - Mark notification(s) as read
export async function PATCH(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const body = await request.json()
    const { notificationId, markAll } = body

    if (markAll) {
      await db.notification.updateMany({
        where: { userId: user.id, isRead: false },
        data: { isRead: true },
      })
      return apiSuccess({ message: 'All notifications marked as read' })
    }

    if (!notificationId) return apiError('Notification ID is required')

    const notification = await db.notification.findUnique({
      where: { id: notificationId },
    })

    if (!notification) return apiError('Notification not found', 404)
    if (notification.userId !== user.id) return apiError('You can only update your own notifications', 403)

    await db.notification.update({
      where: { id: notificationId },
      data: { isRead: true },
    })

    return apiSuccess({ message: 'Notification marked as read' })
  } catch (error: any) {
    console.error('Notification update error:', error)
    return apiError('Failed to update notification', 500)
  }
}
