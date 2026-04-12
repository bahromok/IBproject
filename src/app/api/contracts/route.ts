import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/contracts - List contracts for current user
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {
      OR: [
        { clientId: user.id },
        { freelancerId: user.id },
      ],
    }

    if (status) where.status = status

    const [contracts, total] = await Promise.all([
      db.contract.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          job: {
            select: { id: true, title: true, category: true },
          },
          client: {
            select: {
              id: true,
              clientProfile: {
                select: { displayName: true },
              },
            },
          },
          freelancer: {
            select: {
              id: true,
              freelancerProfile: {
                select: { displayName: true, avatar: true },
              },
            },
          },
          milestones: {
            select: {
              id: true,
              title: true,
              amount: true,
              status: true,
            },
          },
        },
      }),
      db.contract.count({ where }),
    ])

    // Identity protection
    const safeContracts = contracts.map((c) => ({
      ...c,
      client: {
        id: c.client.id,
        displayName: c.client.clientProfile?.displayName || 'Anonymous',
        avatar: null,
      },
      freelancer: {
        id: c.freelancer.id,
        displayName: c.freelancer.freelancerProfile?.displayName || 'Anonymous',
        avatar: c.freelancer.freelancerProfile?.avatar || null,
      },
    }))

    return apiSuccess({
      contracts: safeContracts,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Contracts list error:', error)
    return apiError('Failed to fetch contracts', 500)
  }
}
