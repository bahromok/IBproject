import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyProposalStatus, notifyContractCreated } from '@/lib/notify'

// PUT /api/proposals/[id] - Update proposal status (shortlist, accept, reject - CLIENT only)
export async function PUT(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: proposalId } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const proposal = await db.proposal.findUnique({
      where: { id: proposalId },
      include: { job: true },
    })

    if (!proposal) return apiError('Proposal not found', 404)
    if (proposal.job.clientId !== user.id && user.role !== 'ADMIN') {
      return apiError('Only the job owner can update proposal status', 403)
    }

    const body = await request.json()
    const { status } = body

    if (!['SHORTLISTED', 'ACCEPTED', 'REJECTED'].includes(status)) {
      return apiError('Invalid status. Use SHORTLISTED, ACCEPTED, or REJECTED')
    }

    const updatedProposal = await db.proposal.update({
      where: { id: proposalId },
      data: { status },
    })

    // If accepted, create a contract
    if (status === 'ACCEPTED') {
      // Reject all other pending proposals for this job
      await db.proposal.updateMany({
        where: {
          jobId: proposal.jobId,
          status: 'PENDING',
          id: { not: proposalId },
        },
        data: { status: 'REJECTED' },
      })

      // Update job status
      await db.job.update({
        where: { id: proposal.jobId },
        data: { status: 'IN_PROGRESS' },
      })

      // Create contract
      const contract = await db.contract.create({
        data: {
          jobId: proposal.jobId,
          clientId: proposal.job.clientId,
          freelancerId: proposal.freelancerId,
          proposalId: proposalId,
          title: proposal.job.title,
          description: proposal.job.description,
          contractType: proposal.job.budgetType,
          totalAmount: proposal.proposedAmount,
          startDate: new Date(),
        },
      })

      // Notify freelancer about acceptance (fire-and-forget)
      notifyProposalStatus(proposal.freelancerId, proposalId, proposal.job.title, true)

      // Notify both parties about contract creation (fire-and-forget)
      notifyContractCreated(proposal.freelancerId, contract.id, contract.title)
      notifyContractCreated(proposal.job.clientId, contract.id, contract.title)
    }

    // Notify freelancer about rejection (fire-and-forget)
    if (status === 'REJECTED') {
      notifyProposalStatus(proposal.freelancerId, proposalId, proposal.job.title, false)
    }

    return apiSuccess({ proposal: updatedProposal })
  } catch (error: any) {
    console.error('Proposal update error:', error)
    return apiError('Failed to update proposal', 500)
  }
}

// PATCH /api/proposals/[id] - Withdraw proposal (FREELANCER who owns it only)
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: proposalId } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'FREELANCER') return apiError('Only freelancers can withdraw proposals', 403)

    const proposal = await db.proposal.findUnique({
      where: { id: proposalId },
    })

    if (!proposal) return apiError('Proposal not found', 404)
    if (proposal.freelancerId !== user.id) return apiError('You can only withdraw your own proposals', 403)
    if (proposal.status === 'WITHDRAWN' || proposal.status === 'ACCEPTED') {
      return apiError('Cannot withdraw this proposal', 400)
    }

    const updatedProposal = await db.proposal.update({
      where: { id: proposalId },
      data: { status: 'WITHDRAWN' },
    })

    // Update job proposal count
    await db.job.update({
      where: { id: proposal.jobId },
      data: { proposalsCount: { decrement: 1 } },
    })

    return apiSuccess({ proposal: updatedProposal })
  } catch (error: any) {
    console.error('Proposal withdraw error:', error)
    return apiError('Failed to withdraw proposal', 500)
  }
}
