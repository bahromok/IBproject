import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/contracts/[id]/milestones - List milestones for a contract
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: contractId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const contract = await db.contract.findUnique({ where: { id: contractId } })
    if (!contract) return apiError('Contract not found', 404)

    if (contract.clientId !== user.id && contract.freelancerId !== user.id && user.role !== 'ADMIN') {
      return apiError('You do not have permission to view this contract', 403)
    }

    const milestones = await db.milestone.findMany({
      where: { contractId },
      orderBy: { createdAt: 'asc' },
    })

    return apiSuccess({ milestones })
  } catch (error: any) {
    console.error('Milestones list error:', error)
    return apiError('Failed to fetch milestones', 500)
  }
}

// POST /api/contracts/[id]/milestones - Create milestone (CLIENT only)
export async function POST(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id: contractId } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    if (user.role !== 'CLIENT' && user.role !== 'ADMIN') {
      return apiError('Only clients can create milestones', 403)
    }

    const contract = await db.contract.findUnique({ where: { id: contractId } })
    if (!contract) return apiError('Contract not found', 404)
    if (contract.clientId !== user.id && user.role !== 'ADMIN') {
      return apiError('You can only create milestones for your own contracts', 403)
    }
    if (contract.status !== 'ACTIVE') {
      return apiError('Can only create milestones for active contracts', 400)
    }

    const body = await request.json()
    const { title, description, amount, dueDate } = body

    if (!title || !description || !amount) {
      return apiError('Title, description, and amount are required')
    }

    const milestone = await db.milestone.create({
      data: {
        contractId,
        title,
        description,
        amount: parseFloat(amount),
        dueDate: dueDate ? new Date(dueDate) : null,
      },
    })

    // Notify freelancer
    await db.notification.create({
      data: {
        userId: contract.freelancerId,
        type: 'SYSTEM',
        title: 'New Milestone Added',
        message: `A new milestone "${title}" has been added to "${contract.title}"`,
        actionUrl: `/contracts/${contractId}`,
      },
    })

    return apiSuccess({ milestone }, 201)
  } catch (error: any) {
    console.error('Milestone creation error:', error)
    return apiError('Failed to create milestone', 500)
  }
}
