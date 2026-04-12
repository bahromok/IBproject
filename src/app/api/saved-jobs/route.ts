import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/saved-jobs - Get saved jobs for current user
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const savedJobs = await db.savedJob.findMany({
      where: { userId: user.id },
      orderBy: { createdAt: 'desc' },
      include: {
        job: {
          include: {
            client: {
              select: {
                id: true,
                clientProfile: {
                  select: { displayName: true, location: true },
                },
              },
            },
          },
        },
      },
    })

    const safeJobs = savedJobs.map((sj) => ({
      ...sj,
      job: {
        ...sj.job,
        client: {
          id: sj.job.client.id,
          displayName: sj.job.client.clientProfile?.displayName || 'Anonymous',
          avatar: null,
          location: sj.job.client.clientProfile?.location || null,
        },
      },
    }))

    return apiSuccess({ savedJobs: safeJobs })
  } catch (error: any) {
    console.error('Saved jobs fetch error:', error)
    return apiError('Failed to fetch saved jobs', 500)
  }
}

// POST /api/saved-jobs - Save a job
export async function POST(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const body = await request.json()
    const { jobId } = body

    if (!jobId) return apiError('Job ID is required')

    const job = await db.job.findUnique({ where: { id: jobId } })
    if (!job) return apiError('Job not found', 404)

    const existing = await db.savedJob.findUnique({
      where: { userId_jobId: { userId: user.id, jobId } },
    })
    if (existing) return apiError('Job already saved', 400)

    const savedJob = await db.savedJob.create({
      data: { userId: user.id, jobId },
    })

    return apiSuccess({ savedJob }, 201)
  } catch (error: any) {
    console.error('Save job error:', error)
    return apiError('Failed to save job', 500)
  }
}

// DELETE /api/saved-jobs - Remove saved job
export async function DELETE(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const jobId = searchParams.get('jobId')

    if (!jobId) return apiError('Job ID is required')

    const savedJob = await db.savedJob.findUnique({
      where: { userId_jobId: { userId: user.id, jobId } },
    })
    if (!savedJob) return apiError('Saved job not found', 404)

    await db.savedJob.delete({
      where: { id: savedJob.id },
    })

    return apiSuccess({ message: 'Job removed from saved' })
  } catch (error: any) {
    console.error('Remove saved job error:', error)
    return apiError('Failed to remove saved job', 500)
  }
}
