import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/proposals - List proposals for the current freelancer
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = { freelancerId: user.id }
    if (status) where.status = status

    const [proposals, total] = await Promise.all([
      db.proposal.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          job: {
            select: {
              id: true,
              title: true,
              status: true,
              budgetMin: true,
              budgetMax: true,
              budgetType: true,
              client: {
                select: {
                  id: true,
                  clientProfile: {
                    select: { displayName: true, rating: true },
                  },
                },
              },
            },
          },
        },
      }),
      db.proposal.count({ where }),
    ])

    const safeProposals = proposals.map((p) => ({
      id: p.id,
      proposedAmount: p.proposedAmount,
      proposedDuration: p.proposedDuration,
      coverLetter: p.coverLetter,
      status: p.status,
      createdAt: p.createdAt,
      updatedAt: p.updatedAt,
      job: {
        id: p.job.id,
        title: p.job.title,
        status: p.job.status,
        budgetMin: p.job.budgetMin,
        budgetMax: p.job.budgetMax,
        budgetType: p.job.budgetType,
        client: {
          id: p.job.client.id,
          displayName: p.job.client.clientProfile?.displayName || 'Anonymous',
          avatar: null,
          rating: p.job.client.clientProfile?.rating || 0,
        },
      },
    }))

    return apiSuccess({
      proposals: safeProposals,
      summary: {
        total,
        pending: proposals.filter((p) => p.status === 'PENDING').length,
        accepted: proposals.filter((p) => p.status === 'ACCEPTED').length,
        rejected: proposals.filter((p) => p.status === 'REJECTED' || p.status === 'WITHDRAWN').length,
      },
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Proposals fetch error:', error)
    return apiError('Failed to fetch proposals', 500)
  }
}
