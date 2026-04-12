import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/contracts/[id] - Contract details with milestones
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const contract = await db.contract.findUnique({
      where: { id },
      include: {
        job: {
          select: { id: true, title: true, category: true, description: true },
        },
        client: {
          select: {
            id: true,
            clientProfile: {
              select: { displayName: true, location: true },
            },
          },
        },
        freelancer: {
          select: {
            id: true,
            freelancerProfile: {
              select: { displayName: true, avatar: true, title: true, location: true },
            },
          },
        },
        milestones: {
          orderBy: { createdAt: 'asc' },
        },
        transactions: {
          orderBy: { createdAt: 'desc' },
          select: {
            id: true,
            type: true,
            amount: true,
            status: true,
            description: true,
            createdAt: true,
          },
        },
        reviews: true,
      },
    })

    if (!contract) return apiError('Contract not found', 404)

    // Only involved parties or admin can view
    if (contract.clientId !== user.id && contract.freelancerId !== user.id && user.role !== 'ADMIN') {
      return apiError('You do not have permission to view this contract', 403)
    }

    // Identity protection
    const safeContract = {
      ...contract,
      client: {
        id: contract.client.id,
        displayName: contract.client.clientProfile?.displayName || 'Anonymous',
        avatar: null,
        location: contract.client.clientProfile?.location || null,
      },
      freelancer: {
        id: contract.freelancer.id,
        displayName: contract.freelancer.freelancerProfile?.displayName || 'Anonymous',
        avatar: contract.freelancer.freelancerProfile?.avatar || null,
        title: contract.freelancer.freelancerProfile?.title || null,
        location: contract.freelancer.freelancerProfile?.location || null,
      },
    }

    return apiSuccess({ contract: safeContract })
  } catch (error: any) {
    console.error('Contract detail error:', error)
    return apiError('Failed to fetch contract', 500)
  }
}

// PUT /api/contracts/[id] - Update contract status
export async function PUT(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const contract = await db.contract.findUnique({ where: { id } })
    if (!contract) return apiError('Contract not found', 404)

    if (contract.clientId !== user.id && contract.freelancerId !== user.id && user.role !== 'ADMIN') {
      return apiError('You do not have permission to update this contract', 403)
    }

    const body = await request.json()
    const { status } = body

    const validStatuses = ['ACTIVE', 'PAUSED', 'COMPLETED', 'CANCELLED', 'DISPUTED']
    if (!validStatuses.includes(status)) {
      return apiError(`Invalid status. Use: ${validStatuses.join(', ')}`)
    }

    if (status === 'COMPLETED') {
      // Only client or admin can complete
      if (contract.clientId !== user.id && user.role !== 'ADMIN') {
        return apiError('Only the client can complete the contract', 403)
      }

      const updatedContract = await db.contract.update({
        where: { id },
        data: {
          status: 'COMPLETED',
          endDate: new Date(),
        },
      })

      // Update freelancer stats
      await db.freelancerProfile.update({
        where: { userId: contract.freelancerId },
        data: { completedJobs: { increment: 1 } },
      })

      // Update job status
      await db.job.update({
        where: { id: contract.jobId },
        data: { status: 'COMPLETED' },
      })

      // Notify freelancer
      await db.notification.create({
        data: {
          userId: contract.freelancerId,
          type: 'CONTRACT_COMPLETED',
          title: 'Contract Completed',
          message: `The contract "${contract.title}" has been marked as completed`,
          actionUrl: `/contracts/${id}`,
        },
      })

      return apiSuccess({ contract: updatedContract })
    }

    if (status === 'DISPUTED') {
      const updatedContract = await db.contract.update({
        where: { id },
        data: { status: 'DISPUTED' },
      })

      // Notify the other party
      const otherUserId = contract.clientId === user.id ? contract.freelancerId : contract.clientId
      await db.notification.create({
        data: {
          userId: otherUserId,
          type: 'SYSTEM',
          title: 'Contract Disputed',
          message: `The contract "${contract.title}" has been flagged for dispute`,
          actionUrl: `/contracts/${id}`,
        },
      })

      return apiSuccess({ contract: updatedContract })
    }

    const updatedContract = await db.contract.update({
      where: { id },
      data: { status },
    })

    return apiSuccess({ contract: updatedContract })
  } catch (error: any) {
    console.error('Contract update error:', error)
    return apiError('Failed to update contract', 500)
  }
}
