import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/admin/jobs - List all jobs (with filters)
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const category = searchParams.get('category')
    const isUrgent = searchParams.get('isUrgent')
    const search = searchParams.get('search')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {}

    if (status) where.status = status
    if (category) where.category = category
    if (isUrgent !== null) where.isUrgent = isUrgent === 'true'
    if (search) {
      where.OR = [
        { title: { contains: search } },
        { description: { contains: search } },
      ]
    }

    const [jobs, total] = await Promise.all([
      db.job.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          client: {
            select: {
              id: true,
              email: true,
              clientProfile: { select: { displayName: true } },
            },
          },
          _count: { select: { proposals: true } },
        },
      }),
      db.job.count({ where }),
    ])

    return apiSuccess({
      jobs,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Admin jobs list error:', error)
    return apiError('Failed to fetch jobs', 500)
  }
}

// PATCH /api/admin/jobs - Feature/unfeature job
export async function PATCH(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const body = await request.json()
    const { jobId, action } = body

    if (!jobId || !action) return apiError('Job ID and action are required')

    const job = await db.job.findUnique({ where: { id: jobId } })
    if (!job) return apiError('Job not found', 404)

    if (action === 'feature') {
      const updated = await db.job.update({
        where: { id: jobId },
        data: { isUrgent: true },
      })
      return apiSuccess({ job: updated })
    }

    if (action === 'unfeature') {
      const updated = await db.job.update({
        where: { id: jobId },
        data: { isUrgent: false },
      })
      return apiSuccess({ job: updated })
    }

    return apiError('Invalid action. Use feature or unfeature')
  } catch (error: any) {
    console.error('Admin job update error:', error)
    return apiError('Failed to update job', 500)
  }
}

// DELETE /api/admin/jobs - Remove job
export async function DELETE(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const jobId = searchParams.get('jobId')

    if (!jobId) return apiError('Job ID is required')

    const job = await db.job.findUnique({ where: { id: jobId } })
    if (!job) return apiError('Job not found', 404)

    await db.job.update({
      where: { id: jobId },
      data: { status: 'deleted' },
    })

    return apiSuccess({ message: 'Job removed successfully' })
  } catch (error: any) {
    console.error('Admin job delete error:', error)
    return apiError('Failed to delete job', 500)
  }
}
