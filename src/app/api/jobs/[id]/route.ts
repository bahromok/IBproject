import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/jobs/[id] - Get job details
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params

    const job = await db.job.findUnique({
      where: { id },
      include: {
        client: {
          select: {
            id: true,
            clientProfile: {
              select: {
                displayName: true,
                bio: true,
                companyName: true,
                location: true,
                isVerified: true,
                rating: true,
                totalReviews: true,
                jobsPosted: true,
                createdAt: true,
              },
            },
          },
        },
      },
    })

    if (!job) {
      return apiError('Job not found', 404)
    }

    // Increment view count
    await db.job.update({
      where: { id },
      data: { viewCount: { increment: 1 } },
    })

    // Parse skills from JSON string
    const parsedSkills = typeof job.skills === 'string' ? JSON.parse(job.skills) : job.skills
    const parsedAttachments = typeof job.attachments === 'string' ? JSON.parse(job.attachments) : job.attachments

    // Identity protection
    const safeJob = {
      id: job.id,
      title: job.title,
      description: job.description,
      category: job.category,
      budgetType: job.budgetType,
      budgetMin: job.budgetMin,
      budgetMax: job.budgetMax,
      duration: job.duration,
      experienceLevel: job.experienceLevel,
      skills: parsedSkills,
      status: job.status,
      proposalsCount: job.proposalsCount,
      isUrgent: job.isUrgent,
      attachments: parsedAttachments,
      viewCount: job.viewCount,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
      client: {
        id: job.client.id,
        displayName: job.client.clientProfile?.displayName || 'Anonymous',
        avatar: null,
        bio: job.client.clientProfile?.bio || null,
        companyName: job.client.clientProfile?.companyName || null,
        location: job.client.clientProfile?.location || null,
        isVerified: job.client.clientProfile?.isVerified || false,
        rating: job.client.clientProfile?.rating || 0,
        totalReviews: job.client.clientProfile?.totalReviews || 0,
        jobsPosted: job.client.clientProfile?.jobsPosted || 0,
        memberSince: job.client.clientProfile?.createdAt || null,
      },
    }

    return apiSuccess({ job: safeJob })
  } catch (error: any) {
    console.error('Job detail error:', error)
    return apiError('Failed to fetch job', 500)
  }
}

// PUT /api/jobs/[id] - Update job (CLIENT who owns it only)
export async function PUT(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'CLIENT') return apiError('Only clients can update jobs', 403)

    const existing = await db.job.findUnique({ where: { id } })
    if (!existing) return apiError('Job not found', 404)
    if (existing.clientId !== user.id) return apiError('You can only update your own jobs', 403)
    if (existing.status !== 'open') return apiError('Can only update open jobs', 400)

    const body = await request.json()
    const { title, description, category, skills, budgetType, budgetMin, budgetMax, duration, experienceLevel, isUrgent } = body

    const job = await db.job.update({
      where: { id },
      data: {
        ...(title !== undefined && { title }),
        ...(description !== undefined && { description }),
        ...(category !== undefined && { category }),
        ...(skills !== undefined && { skills: JSON.stringify(skills) }),
        ...(budgetType !== undefined && { budgetType }),
        ...(budgetMin !== undefined && { budgetMin: parseFloat(budgetMin) || null }),
        ...(budgetMax !== undefined && { budgetMax: parseFloat(budgetMax) || null }),
        ...(duration !== undefined && { duration }),
        ...(experienceLevel !== undefined && { experienceLevel }),
        ...(isUrgent !== undefined && { isUrgent }),
      },
    })

    return apiSuccess({ job })
  } catch (error: any) {
    console.error('Job update error:', error)
    return apiError('Failed to update job', 500)
  }
}

// DELETE /api/jobs/[id] - Soft delete job (CLIENT who owns it only)
export async function DELETE(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)

    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'CLIENT') return apiError('Only clients can delete jobs', 403)

    const existing = await db.job.findUnique({ where: { id } })
    if (!existing) return apiError('Job not found', 404)
    if (existing.clientId !== user.id) return apiError('You can only delete your own jobs', 403)

    const job = await db.job.update({
      where: { id },
      data: { status: 'deleted' },
    })

    return apiSuccess({ message: 'Job deleted successfully' })
  } catch (error: any) {
    console.error('Job delete error:', error)
    return apiError('Failed to delete job', 500)
  }
}
