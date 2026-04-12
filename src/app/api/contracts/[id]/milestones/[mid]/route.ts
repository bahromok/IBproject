import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyMilestoneUpdate, notifyPayment } from '@/lib/notify'

// PUT /api/contracts/[id]/milestones/[mid] - Update milestone
export async function PUT(
  request: NextRequest,
  { params }: { params: Promise<{ id: string; mid: string }> }
) {
  try {
    const { id: contractId, mid: milestoneId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const milestone = await db.milestone.findUnique({
      where: { id: milestoneId },
      include: { contract: true },
    })

    if (!milestone) return apiError('Milestone not found', 404)
    if (milestone.contractId !== contractId) return apiError('Milestone does not belong to this contract', 400)

    const isClient = milestone.contract.clientId === user.id
    const isFreelancer = milestone.contract.freelancerId === user.id
    if (!isClient && !isFreelancer && user.role !== 'ADMIN') {
      return apiError('You do not have permission to update this milestone', 403)
    }

    const body = await request.json()
    const { title, description, amount, dueDate, status, submissionNote } = body

    // Freelancer can submit work (change to SUBMITTED)
    if (status === 'SUBMITTED' && isFreelancer) {
      const updated = await db.milestone.update({
        where: { id: milestoneId },
        data: {
          status: 'SUBMITTED',
          submissionNotes: submissionNote || null,
          submittedAt: new Date(),
        },
      })

      // Notify client about submission (fire-and-forget)
      notifyMilestoneUpdate(milestone.contract.clientId, contractId, milestone.title, 'SUBMITTED')

      return apiSuccess({ milestone: updated })
    }

    // Freelancer can start working on milestone
    if (status === 'IN_PROGRESS' && isFreelancer) {
      const updated = await db.milestone.update({
        where: { id: milestoneId },
        data: { status: 'IN_PROGRESS' },
      })
      return apiSuccess({ milestone: updated })
    }

    // Client can update milestone details
    if (isClient || user.role === 'ADMIN') {
      const updateData: any = {}
      if (title !== undefined) updateData.title = title
      if (description !== undefined) updateData.description = description
      if (amount !== undefined) updateData.amount = parseFloat(amount)
      if (dueDate !== undefined) updateData.dueDate = dueDate ? new Date(dueDate) : null
      if (status !== undefined && ['PENDING', 'IN_PROGRESS'].includes(status)) {
        updateData.status = status
      }

      const updated = await db.milestone.update({
        where: { id: milestoneId },
        data: updateData,
      })

      return apiSuccess({ milestone: updated })
    }

    return apiError('Invalid update operation', 400)
  } catch (error: any) {
    console.error('Milestone update error:', error)
    return apiError('Failed to update milestone', 500)
  }
}

// PATCH /api/contracts/[id]/milestones/[mid] - Approve/reject milestone (CLIENT only)
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string; mid: string }> }
) {
  try {
    const { id: contractId, mid: milestoneId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const milestone = await db.milestone.findUnique({
      where: { id: milestoneId },
      include: { contract: true },
    })

    if (!milestone) return apiError('Milestone not found', 404)
    if (milestone.contractId !== contractId) return apiError('Milestone does not belong to this contract', 400)

    if (milestone.contract.clientId !== user.id && user.role !== 'ADMIN') {
      return apiError('Only the client can approve or reject milestones', 403)
    }

    const body = await request.json()
    const { action, feedback } = body // action: 'approve' or 'reject'

    if (!['approve', 'reject'].includes(action)) {
      return apiError('Action must be "approve" or "reject"')
    }

    if (milestone.status !== 'SUBMITTED') {
      return apiError('Can only approve or reject submitted milestones', 400)
    }

    const platformFee = milestone.amount * 0.1
    const netAmount = milestone.amount - platformFee

    if (action === 'approve') {
      const updated = await db.milestone.update({
        where: { id: milestoneId },
        data: {
          status: 'APPROVED',
          approvalNotes: feedback || null,
          approvedAt: new Date(),
        },
      })

      // Create payment release transaction (client → freelancer)
      await db.transaction.create({
        data: {
          fromUserId: milestone.contract.clientId,
          toUserId: milestone.contract.freelancerId,
          contractId,
          milestoneId,
          type: 'PAYMENT_RELEASE',
          amount: milestone.amount,
          platformFee,
          netAmount,
          status: 'COMPLETED',
          description: `Milestone "${milestone.title}" approved — payment released`,
        },
      })

      // Notify freelancer about approval + payment (fire-and-forget)
      notifyMilestoneUpdate(milestone.contract.freelancerId, contractId, milestone.title, 'APPROVED')
      notifyPayment(milestone.contract.freelancerId, contractId, netAmount)

      return apiSuccess({ milestone: updated })
    }

    if (action === 'reject') {
      const updated = await db.milestone.update({
        where: { id: milestoneId },
        data: {
          status: 'REJECTED',
          approvalNotes: feedback || null,
        },
      })

      // Notify freelancer about rejection (fire-and-forget)
      notifyMilestoneUpdate(milestone.contract.freelancerId, contractId, milestone.title, 'REJECTED')

      return apiSuccess({ milestone: updated })
    }

    return apiError('Invalid action', 400)
  } catch (error: any) {
    console.error('Milestone approval error:', error)
    return apiError('Failed to process milestone', 500)
  }
}
