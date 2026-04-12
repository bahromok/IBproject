import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyReviewReceived } from '@/lib/notify'

// POST /api/reviews - Create review (after contract completion)
export async function POST(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const body = await request.json()
    const { contractId, rating, comment, quality, communication, professionalism } = body

    if (!contractId || !rating) return apiError('Contract ID and rating are required')
    if (rating < 1 || rating > 5) return apiError('Rating must be between 1 and 5')

    const contract = await db.contract.findUnique({
      where: { id: contractId },
      include: { reviews: true },
    })

    if (!contract) return apiError('Contract not found', 404)
    if (contract.status !== 'COMPLETED') return apiError('Can only review completed contracts', 400)

    // Determine the other party
    let revieweeId: string
    if (contract.clientId === user.id) {
      revieweeId = contract.freelancerId
    } else if (contract.freelancerId === user.id) {
      revieweeId = contract.clientId
    } else {
      return apiError('You are not part of this contract', 403)
    }

    // Check if already reviewed
    const existingReview = contract.reviews.find((r) => r.reviewerId === user.id)
    if (existingReview) return apiError('You have already reviewed this contract', 400)

    const review = await db.review.create({
      data: {
        contractId,
        reviewerId: user.id,
        revieweeId,
        rating,
        quality: quality || rating,
        communication: communication || rating,
        professionalism: professionalism || rating,
        comment: comment || null,
      },
    })

    // Update target's rating (if freelancer)
    if (contract.freelancerId === revieweeId) {
      const allReviews = await db.review.findMany({
        where: { revieweeId },
      })
      const avgRating = allReviews.reduce((sum, r) => sum + r.rating, 0) / allReviews.length

      await db.freelancerProfile.update({
        where: { userId: revieweeId },
        data: {
          rating: Math.round(avgRating * 10) / 10,
          totalReviews: allReviews.length,
        },
      })
    }

    // Notify the reviewed user (fire-and-forget)
    notifyReviewReceived(revieweeId, contractId, rating)

    return apiSuccess({ review }, 201)
  } catch (error: any) {
    console.error('Review creation error:', error)
    return apiError('Failed to create review', 500)
  }
}

// GET /api/reviews - Get reviews for a user
export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url)
    const userId = searchParams.get('userId')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    if (!userId) return apiError('User ID is required')

    const where: any = { revieweeId: userId }

    const [reviews, total] = await Promise.all([
      db.review.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
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
      }),
      db.review.count({ where }),
    ])

    // Identity protection
    const safeReviews = reviews.map((r) => ({
      ...r,
      reviewer: {
        id: r.reviewer.id,
        displayName:
          r.reviewer.clientProfile?.displayName ||
          r.reviewer.freelancerProfile?.displayName ||
          'Anonymous',
        avatar:
          r.reviewer.freelancerProfile?.avatar ||
          null,
      },
    }))

    // Calculate average
    const allReviewsForUser = await db.review.findMany({
      where: { revieweeId: userId },
      select: { rating: true },
    })
    const avgRating = allReviewsForUser.length
      ? allReviewsForUser.reduce((sum, r) => sum + r.rating, 0) / allReviewsForUser.length
      : 0

    return apiSuccess({
      reviews: safeReviews,
      stats: {
        total: allReviewsForUser.length,
        averageRating: Math.round(avgRating * 10) / 10,
      },
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Reviews fetch error:', error)
    return apiError('Failed to fetch reviews', 500)
  }
}
