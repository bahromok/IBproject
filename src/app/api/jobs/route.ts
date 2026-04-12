import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess } from '@/lib/api-auth'
import { checkAndFlagUser } from '@/lib/fraud-detection'

// GET /api/jobs - List jobs with filters
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const category = searchParams.get('category')
    const budgetMin = searchParams.get('budgetMin')
    const budgetMax = searchParams.get('budgetMax')
    const clientId = searchParams.get('clientId')
    // When clientId is provided, don't default status to 'open' (client sees all their jobs)
    const statusParam = searchParams.get('status') || (clientId ? undefined : 'OPEN')
    const search = searchParams.get('search')
    const budgetType = searchParams.get('budgetType')
    const experienceLevel = searchParams.get('experienceLevel')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')
    const sortBy = searchParams.get('sortBy') || 'createdAt'
    const sortOrder = searchParams.get('sortOrder') || 'desc'

    const skip = (page - 1) * limit

    const where: any = {}
    if (statusParam) where.status = statusParam
    if (clientId) where.clientId = clientId

    if (category) where.category = category
    if (budgetType) where.budgetType = budgetType
    if (experienceLevel) where.experienceLevel = experienceLevel
    if (budgetMin) where.budgetMin = { gte: parseFloat(budgetMin) }
    if (budgetMax) where.budgetMax = { lte: parseFloat(budgetMax) }
    if (search) {
      where.OR = [
        { title: { contains: search } },
        { description: { contains: search } },
        { skills: { contains: search } },
      ]
    }

    const orderBy: any = {}
    if (sortBy === 'budget') {
      orderBy.budgetMax = sortOrder === 'asc' ? 'asc' : 'desc'
    } else if (sortBy === 'proposalsCount') {
      orderBy.proposalsCount = sortOrder === 'asc' ? 'asc' : 'desc'
    } else {
      orderBy[sortBy] = sortOrder === 'asc' ? 'asc' : 'desc'
    }

    const [jobs, total] = await Promise.all([
      db.job.findMany({
        where,
        skip,
        take: limit,
        orderBy,
        include: {
          client: {
            select: {
              id: true,
              clientProfile: {
                select: {
                  displayName: true,
                  location: true,
                  isVerified: true,
                  rating: true,
                },
              },
            },
          },
        },
      }),
      db.job.count({ where }),
    ])

    // Identity protection: only return displayName, never email or real name
    const safeJobs = jobs.map((job) => ({
      id: job.id,
      title: job.title,
      description: job.description,
      category: job.category,
      budgetType: job.budgetType,
      budgetMin: job.budgetMin,
      budgetMax: job.budgetMax,
      duration: job.duration,
      experienceLevel: job.experienceLevel,
      skills: (() => { try { return typeof job.skills === 'string' ? JSON.parse(job.skills) : job.skills } catch { return typeof job.skills === 'string' ? job.skills.split(',').map((s: string) => s.trim()).filter(Boolean) : job.skills } })(),
      status: job.status,
      proposalsCount: job.proposalsCount,
      isUrgent: job.isUrgent,
      viewCount: job.viewCount,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
      client: {
        id: job.client.id,
        displayName: job.client.clientProfile?.displayName || 'Anonymous',
        avatar: null,
        location: job.client.clientProfile?.location || null,
        isVerified: job.client.clientProfile?.isVerified || false,
        rating: job.client.clientProfile?.rating || 0,
      },
    }))

    return apiSuccess({
      jobs: safeJobs,
      pagination: {
        page,
        limit,
        total,
        totalPages: Math.ceil(total / limit),
      },
    })
  } catch (error: any) {
    console.error('Jobs list error:', error)
    return apiError('Failed to fetch jobs', 500)
  }
}

// POST /api/jobs - Create job (CLIENT only)
export async function POST(request: NextRequest) {
  try {
    const { getAuthenticatedUser, isAuthUser } = await import('@/lib/api-auth')
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    if (authResult.role !== 'CLIENT') return apiError('Only clients can create jobs', 403)
    const clientId = authResult.id

    const body = await request.json()
    const { title, description, category, skills, budgetType, budgetMin, budgetMax, duration, experienceLevel, isUrgent, status } = body

    if (!title || !description || !category) {
      return apiError('Title, description, and category are required')
    }

    if (budgetType === 'FIXED_PRICE' && !budgetMin) {
      return apiError('Budget is required for fixed-price jobs')
    }

    const job = await db.job.create({
      data: {
        clientId,
        title,
        description,
        category,
        skills: JSON.stringify(skills || []),
        budgetType: budgetType || 'FIXED_PRICE',
        budgetMin: budgetMin ? parseFloat(budgetMin) : null,
        budgetMax: budgetMax ? parseFloat(budgetMax) : null,
        duration: duration || null,
        experienceLevel: experienceLevel || 'INTERMEDIATE',
        isUrgent: isUrgent || false,
        status: status || 'OPEN',
      },
    })

    // Fire-and-forget: check if client shows suspicious activity (rapid job posting)
    checkAndFlagUser(clientId).catch(() => { })

    return apiSuccess({ job }, 201)
  } catch (error: any) {
    console.error('Job creation error:', error)
    return apiError('Failed to create job', 500)
  }
}
