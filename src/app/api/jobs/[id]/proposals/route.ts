import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyNewProposal } from '@/lib/notify'
import { checkAndFlagUser } from '@/lib/fraud-detection'

// GET /api/jobs/[id]/proposals - List proposals for a job (CLIENT who owns job only)
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: jobId } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const job = await db.job.findUnique({ where: { id: jobId } })
    if (!job) return apiError('Job not found', 404)

    // Only the job owner or admin can see proposals
    if (job.clientId !== user.id && user.role !== 'ADMIN') {
      return apiError('Only the job owner can view proposals', 403)
    }

    const proposals = await db.proposal.findMany({
      where: { jobId },
      include: {
        freelancer: {
          select: {
            id: true,
            freelancerProfile: {
              select: {
                displayName: true,
                avatar: true,
                title: true,
                rating: true,
                completedJobs: true,
                location: true,
                skills: true,
              },
            },
          },
        },
      },
      orderBy: { createdAt: 'desc' },
    })

    const safeProposals = proposals.map((p) => ({
      id: p.id,
      coverLetter: p.coverLetter,
      proposedAmount: p.proposedAmount,
      proposedDuration: p.proposedDuration,
      status: p.status,
      createdAt: p.createdAt,
      freelancer: {
        id: p.freelancer.id,
        displayName: p.freelancer.freelancerProfile?.displayName || 'Anonymous',
        avatar: p.freelancer.freelancerProfile?.avatar || null,
        title: p.freelancer.freelancerProfile?.title || null,
        rating: p.freelancer.freelancerProfile?.rating || 0,
        completedJobs: p.freelancer.freelancerProfile?.completedJobs || 0,
        location: p.freelancer.freelancerProfile?.location || null,
        skills: p.freelancer.freelancerProfile?.skills || '[]',
      },
    }))

    return apiSuccess({ proposals: safeProposals })
  } catch (error: any) {
    console.error('Proposals list error:', error)
    return apiError('Failed to fetch proposals', 500)
  }
}

// POST /api/jobs/[id]/proposals - Submit proposal (FREELANCER only)
export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: jobId } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'FREELANCER') return apiError('Only freelancers can submit proposals', 403)

    const job = await db.job.findUnique({ where: { id: jobId } })
    if (!job) return apiError('Job not found', 404)
    if (job.status !== 'open') return apiError('This job is no longer accepting proposals', 400)

    const freelancerProfile = await db.freelancerProfile.findUnique({
      where: { userId: user.id },
    })
    if (!freelancerProfile) return apiError('Freelancer profile not found', 404)

    // Check if already proposed
    const existingProposal = await db.proposal.findUnique({
      where: {
        jobId_freelancerId: { jobId, freelancerId: user.id },
      },
    })
    if (existingProposal) return apiError('You have already submitted a proposal for this job', 400)

    const body = await request.json()
    const { coverLetter, proposedAmount, proposedDuration } = body

    if (!coverLetter || !proposedAmount) {
      return apiError('Cover letter and proposed amount are required')
    }

    const proposal = await db.proposal.create({
      data: {
        jobId,
        freelancerId: user.id,
        coverLetter,
        proposedAmount: parseFloat(proposedAmount),
        proposedDuration: proposedDuration || null,
      },
    })

    // Update proposal count
    await db.job.update({
      where: { id: jobId },
      data: { proposalsCount: { increment: 1 } },
    })

    // Notify client about new proposal (fire-and-forget)
    notifyNewProposal(job.clientId, jobId, freelancerProfile.displayName, job.title)

    // Fire-and-forget: check if freelancer shows suspicious activity
    checkAndFlagUser(user.id).catch(() => { })

    return apiSuccess({ proposal }, 201)
  } catch (error: any) {
    console.error('Proposal creation error:', error)
    return apiError('Failed to submit proposal', 500)
  }
}
