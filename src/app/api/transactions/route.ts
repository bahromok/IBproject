import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/transactions - Get transactions for current user
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')
    const type = searchParams.get('type')
    const status = searchParams.get('status')

    const where: any = {
      OR: [
        { fromUserId: user.id },
        { toUserId: user.id },
      ],
    }
    if (type) where.type = type
    if (status) where.status = status

    const [transactions, total] = await Promise.all([
      db.transaction.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          fromUser: {
            select: {
              id: true,
              clientProfile: { select: { displayName: true } },
              freelancerProfile: { select: { displayName: true } },
            },
          },
          toUser: {
            select: {
              id: true,
              clientProfile: { select: { displayName: true } },
              freelancerProfile: { select: { displayName: true } },
            },
          },
          contract: {
            select: { id: true, title: true },
          },
          milestone: {
            select: { id: true, title: true },
          },
        },
      }),
      db.transaction.count({ where }),
    ])

    // Calculate totals based on user role
    const allUserTransactions = await db.transaction.findMany({
      where: {
        OR: [
          { fromUserId: user.id },
          { toUserId: user.id },
        ],
        status: 'COMPLETED',
      },
    })

    // Total spent: sum of deposits this user made (fromUserId = user)
    const totalSpent = allUserTransactions
      .filter((t) => t.fromUserId === user.id && t.type === 'ESCROW_DEPOSIT')
      .reduce((sum, t) => sum + t.amount, 0)

    // Total earnings: sum of releases TO this user (toUserId = user)
    const totalEarnings = allUserTransactions
      .filter((t) => t.toUserId === user.id && t.type === 'PAYMENT_RELEASE')
      .reduce((sum, t) => sum + t.netAmount, 0)

    // Platform fees
    const totalPlatformFees = allUserTransactions
      .filter((t) => t.type === 'PLATFORM_FEE')
      .reduce((sum, t) => sum + t.amount, 0)

    return apiSuccess({
      transactions,
      summary: {
        totalSpent,
        totalEarnings,
        totalPlatformFees,
        totalTransactions: total,
      },
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Transactions fetch error:', error)
    return apiError('Failed to fetch transactions', 500)
  }
}
