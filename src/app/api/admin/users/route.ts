import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/admin/users - List all users with filters, search, pagination
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const search = searchParams.get('search')
    const role = searchParams.get('role')
    const isVerified = searchParams.get('isVerified')
    const isSuspended = searchParams.get('isSuspended')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {}

    if (search) {
      where.OR = [
        { email: { contains: search } },
        { clientProfile: { displayName: { contains: search } } },
        { freelancerProfile: { displayName: { contains: search } } },
      ]
    }
    if (role) where.role = role
    if (isVerified !== null) where.isVerified = isVerified === 'true'
    if (isSuspended !== null) where.isSuspended = isSuspended === 'true'

    const [users, total] = await Promise.all([
      db.user.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        select: {
          id: true,
          email: true,
          role: true,
          isVerified: true,
          isSuspended: true,
          createdAt: true,
          updatedAt: true,
          clientProfile: { select: { displayName: true, id: true } },
          freelancerProfile: { select: { displayName: true, id: true, rating: true, completedJobs: true } },
          _count: {
            select: {
              jobsPosted: true,
              clientContracts: true,
              freelancerContracts: true,
            },
          },
        },
      }),
      db.user.count({ where }),
    ])

    return apiSuccess({
      users,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Admin users list error:', error)
    return apiError('Failed to fetch users', 500)
  }
}

// PATCH /api/admin/users - Suspend/unsuspend or verify user
export async function PATCH(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const body = await request.json()
    const { userId, action } = body

    if (!userId || !action) return apiError('User ID and action are required')

    const targetUser = await db.user.findUnique({ where: { id: userId } })
    if (!targetUser) return apiError('User not found', 404)

    if (action === 'suspend') {
      const updated = await db.user.update({
        where: { id: userId },
        data: { isSuspended: true },
      })
      return apiSuccess({ user: updated })
    }

    if (action === 'unsuspend') {
      const updated = await db.user.update({
        where: { id: userId },
        data: { isSuspended: false },
      })
      return apiSuccess({ user: updated })
    }

    if (action === 'verify') {
      const updated = await db.user.update({
        where: { id: userId },
        data: { isVerified: true },
      })
      return apiSuccess({ user: updated })
    }

    return apiError('Invalid action. Use suspend, unsuspend, or verify')
  } catch (error: any) {
    console.error('Admin user update error:', error)
    return apiError('Failed to update user', 500)
  }
}
