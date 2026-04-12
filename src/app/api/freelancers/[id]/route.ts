import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess } from '@/lib/api-auth'

// GET /api/freelancers/[id] - Get freelancer public profile
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params

    const profile = await db.freelancerProfile.findUnique({
      where: { id },
      include: {
        user: {
          select: {
            id: true,
            role: true,
            isVerified: true,
          },
        },
      },
    })

    if (!profile) return apiError('Freelancer profile not found', 404)

    // Get reviews for this freelancer
    const reviews = await db.review.findMany({
      where: { revieweeId: profile.userId },
      orderBy: { createdAt: 'desc' },
      take: 10,
      include: {
        reviewer: {
          select: {
            id: true,
            clientProfile: { select: { displayName: true } },
            freelancerProfile: { select: { displayName: true, avatar: true } },
          },
        },
        contract: {
          select: { id: true, title: true },
        },
      },
    })

    // Get active contracts count
    const activeContracts = await db.contract.count({
      where: {
        freelancerId: profile.userId,
        status: 'ACTIVE',
      },
    })

    // Get completed contracts count
    const completedContracts = await db.contract.count({
      where: {
        freelancerId: profile.userId,
        status: 'COMPLETED',
      },
    })

    // Identity protection: NEVER expose real identity
    const safeProfile = {
      id: profile.id,
      displayName: profile.displayName,
      avatar: profile.avatar,
      title: profile.title,
      bio: profile.bio,
      hourlyRate: profile.hourlyRate,
      skills: typeof profile.skills === 'string' ? JSON.parse(profile.skills) : profile.skills,
      portfolio: typeof profile.portfolio === 'string' ? JSON.parse(profile.portfolio) : profile.portfolio,
      category: profile.category,
      location: profile.location,
      timezone: profile.timezone,
      availabilityStatus: profile.availabilityStatus,
      rating: profile.rating,
      totalReviews: profile.totalReviews,
      completedJobs: profile.completedJobs,
      experienceLevel: profile.experienceLevel,
      englishLevel: profile.englishLevel,
      isStudent: profile.isStudent,
      studentInstitution: profile.studentInstitution,
      studentMajor: profile.studentMajor,
      expectedGraduation: profile.expectedGraduation,
      isVerified: profile.user.isVerified,
      isFeatured: profile.isFeatured,
      stats: {
        activeContracts,
        completedContracts,
      },
      reviews: reviews.map((r) => ({
        id: r.id,
        rating: r.rating,
        quality: r.quality,
        communication: r.communication,
        professionalism: r.professionalism,
        comment: r.comment,
        createdAt: r.createdAt,
        reviewer: {
          displayName:
            r.reviewer.clientProfile?.displayName ||
            r.reviewer.freelancerProfile?.displayName ||
            'Anonymous',
          avatar:
            r.reviewer.freelancerProfile?.avatar ||
            null,
        },
        contract: r.contract,
      })),
    }

    return apiSuccess({ freelancer: safeProfile })
  } catch (error: any) {
    console.error('Freelancer profile error:', error)
    return apiError('Failed to fetch freelancer profile', 500)
  }
}
