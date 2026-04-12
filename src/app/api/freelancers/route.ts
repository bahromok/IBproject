import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess } from '@/lib/api-auth'

// GET /api/freelancers - Browse freelancers with search and filters
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const search = searchParams.get('search')
    const category = searchParams.get('category')
    const skills = searchParams.get('skills')
    const availability = searchParams.get('availability')
    const minRating = searchParams.get('minRating')
    const minHourlyRate = searchParams.get('minHourlyRate')
    const maxHourlyRate = searchParams.get('maxHourlyRate')
    const isStudent = searchParams.get('isStudent')
    const isFeatured = searchParams.get('isFeatured')
    const sortBy = searchParams.get('sortBy') || 'rating'
    const sortOrder = searchParams.get('sortOrder') || 'desc'
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {}

    if (search) {
      where.OR = [
        { displayName: { contains: search } },
        { title: { contains: search } },
        { bio: { contains: search } },
        { skills: { contains: search } },
      ]
    }

    if (availability) where.availabilityStatus = availability
    if (minRating) where.rating = { gte: parseFloat(minRating) }
    if (skills) where.skills = { contains: skills }
    if (isStudent === 'true') where.isStudent = true
    if (isFeatured === 'true') where.isFeatured = true
    if (category) where.category = category

    if (minHourlyRate || maxHourlyRate) {
      where.hourlyRate = {}
      if (minHourlyRate) where.hourlyRate.gte = parseFloat(minHourlyRate)
      if (maxHourlyRate) where.hourlyRate.lte = parseFloat(maxHourlyRate)
    }

    const orderBy: any = {}
    if (sortBy === 'rating') {
      orderBy.rating = sortOrder === 'asc' ? 'asc' : 'desc'
    } else if (sortBy === 'completedJobs') {
      orderBy.completedJobs = sortOrder === 'asc' ? 'asc' : 'desc'
    } else if (sortBy === 'hourlyRate') {
      orderBy.hourlyRate = sortOrder === 'asc' ? 'asc' : 'desc'
    } else if (sortBy === 'createdAt') {
      orderBy.createdAt = sortOrder === 'asc' ? 'asc' : 'desc'
    } else {
      orderBy[sortBy] = sortOrder === 'asc' ? 'asc' : 'desc'
    }

    const [freelancers, total] = await Promise.all([
      db.freelancerProfile.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy,
        select: {
          id: true,
          displayName: true,
          avatar: true,
          title: true,
          bio: true,
          hourlyRate: true,
          skills: true,
          category: true,
          location: true,
          timezone: true,
          availabilityStatus: true,
          rating: true,
          totalReviews: true,
          completedJobs: true,
          experienceLevel: true,
          englishLevel: true,
          isStudent: true,
          studentInstitution: true,
          studentMajor: true,
          isFeatured: true,
          createdAt: true,
        },
      }),
      db.freelancerProfile.count({ where }),
    ])

    const safeFreelancers = freelancers.map((f) => {
      let parsedSkills: string[] = []
      if (f.skills) {
        if (typeof f.skills === 'string') {
          try {
            parsedSkills = JSON.parse(f.skills)
          } catch {
            parsedSkills = f.skills.split(',').map((s: string) => s.trim()).filter(Boolean)
          }
        } else if (Array.isArray(f.skills)) {
          parsedSkills = f.skills
        }
      }
      return { ...f, skills: parsedSkills }
    })

    return apiSuccess({
      freelancers: safeFreelancers,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Freelancers browse error:', error)
    return apiError('Failed to fetch freelancers', 500)
  }
}
