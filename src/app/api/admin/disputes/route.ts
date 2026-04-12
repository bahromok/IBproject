import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/admin/disputes - List disputed contracts
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where = { status: 'DISPUTED' as const }

    const [contracts, total] = await Promise.all([
      db.contract.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { updatedAt: 'desc' },
        include: {
          job: { select: { id: true, title: true, category: true } },
          client: {
            select: {
              id: true,
              clientProfile: { select: { displayName: true } },
            },
          },
          freelancer: {
            select: {
              id: true,
              freelancerProfile: { select: { displayName: true } },
            },
          },
          milestones: {
            select: { id: true, title: true, amount: true, status: true },
          },
          _count: { select: { transactions: true } },
        },
      }),
      db.contract.count({ where }),
    ])

    return apiSuccess({
      disputes: contracts,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Admin disputes list error:', error)
    return apiError('Failed to fetch disputes', 500)
  }
}

// PUT /api/admin/disputes - Resolve dispute (release/refund)
export async function PUT(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const body = await request.json()
    const { contractId, resolution } = body // resolution: 'release' or 'refund'

    if (!contractId || !resolution) return apiError('Contract ID and resolution are required')
    if (!['release', 'refund'].includes(resolution)) {
      return apiError('Resolution must be "release" or "refund"')
    }

    const contract = await db.contract.findUnique({
      where: { id: contractId },
      include: { milestones: true },
    })

    if (!contract) return apiError('Contract not found', 404)
    if (contract.status !== 'DISPUTED') return apiError('This contract is not in dispute', 400)

    const platformFee = contract.totalAmount * 0.1
    const netAmount = contract.totalAmount - platformFee

    if (resolution === 'release') {
      // Release funds to freelancer
      await db.contract.update({
        where: { id: contractId },
        data: { status: 'COMPLETED', endDate: new Date() },
      })

      // Create payment release transaction
      await db.transaction.create({
        data: {
          fromUserId: contract.clientId,
          toUserId: contract.freelancerId,
          contractId,
          type: 'PAYMENT_RELEASE',
          amount: contract.totalAmount,
          platformFee,
          netAmount,
          status: 'COMPLETED',
          description: `Dispute resolved: funds released for "${contract.title}"`,
        },
      })

      // Update freelancer stats
      await db.freelancerProfile.update({
        where: { userId: contract.freelancerId },
        data: { completedJobs: { increment: 1 } },
      })

      // Update job
      await db.job.update({
        where: { id: contract.jobId },
        data: { status: 'COMPLETED' },
      })

      // Notify both parties
      await db.notification.create({
        data: {
          userId: contract.clientId,
          type: 'SYSTEM',
          title: 'Dispute Resolved',
          message: `The dispute for "${contract.title}" has been resolved. Funds released.`,
          actionUrl: `/contracts/${contractId}`,
        },
      })
      await db.notification.create({
        data: {
          userId: contract.freelancerId,
          type: 'SYSTEM',
          title: 'Dispute Resolved',
          message: `The dispute for "${contract.title}" has been resolved in your favor.`,
          actionUrl: `/contracts/${contractId}`,
        },
      })
    } else if (resolution === 'refund') {
      // Refund to client
      await db.contract.update({
        where: { id: contractId },
        data: { status: 'CANCELLED', endDate: new Date() },
      })

      // Create refund transaction
      await db.transaction.create({
        data: {
          fromUserId: contract.freelancerId,
          toUserId: contract.clientId,
          contractId,
          type: 'REFUND',
          amount: contract.totalAmount,
          platformFee: 0,
          netAmount: contract.totalAmount,
          status: 'COMPLETED',
          description: `Dispute resolved: refund for "${contract.title}"`,
        },
      })

      // Update job
      await db.job.update({
        where: { id: contract.jobId },
        data: { status: 'CANCELLED' },
      })

      // Notify both parties
      await db.notification.create({
        data: {
          userId: contract.clientId,
          type: 'SYSTEM',
          title: 'Dispute Resolved',
          message: `The dispute for "${contract.title}" has been resolved. Refund issued.`,
          actionUrl: `/contracts/${contractId}`,
        },
      })
      await db.notification.create({
        data: {
          userId: contract.freelancerId,
          type: 'SYSTEM',
          title: 'Dispute Resolved',
          message: `The dispute for "${contract.title}" has been resolved. Contract cancelled.`,
          actionUrl: `/contracts/${contractId}`,
        },
      })
    }

    return apiSuccess({ message: `Dispute resolved: ${resolution}` })
  } catch (error: any) {
    console.error('Admin dispute resolution error:', error)
    return apiError('Failed to resolve dispute', 500)
  }
}
