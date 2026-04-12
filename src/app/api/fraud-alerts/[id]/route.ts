import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// PATCH /api/fraud-alerts/[id] - Update alert status, add admin note (Admin only)
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const alert = await db.fraudAlert.findUnique({ where: { id } })
    if (!alert) return apiError('Fraud alert not found', 404)

    const body = await request.json()
    const { status, adminNote } = body

    const validStatuses = ['PENDING', 'REVIEWED', 'DISMISSED', 'ACTION_TAKEN']
    if (status && !validStatuses.includes(status)) {
      return apiError(`Invalid status. Use: ${validStatuses.join(', ')}`)
    }

    const updated = await db.fraudAlert.update({
      where: { id },
      data: {
        ...(status && { status }),
        ...(adminNote !== undefined && { adminNote }),
        ...(status && status !== 'PENDING'
          ? { reviewedBy: user.id, reviewedAt: new Date() }
          : {}),
      },
    })

    return apiSuccess({ alert: updated })
  } catch (error: any) {
    console.error('Fraud alert update error:', error)
    return apiError('Failed to update fraud alert', 500)
  }
}

// GET /api/fraud-alerts/[id] - Get single alert with user details (Admin only)
export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const alert = await db.fraudAlert.findUnique({
      where: { id },
      include: {
        user: {
          select: {
            id: true,
            email: true,
            role: true,
            isSuspended: true,
            isVerified: true,
            createdAt: true,
            freelancerProfile: { select: { displayName: true, rating: true, completedJobs: true } },
            clientProfile: { select: { displayName: true, rating: true, totalSpent: true } },
          },
        },
      },
    })

    if (!alert) return apiError('Fraud alert not found', 404)

    // Fetch user's report history
    const [reportsReceived, reportsFiled] = await Promise.all([
      db.report.findMany({
        where: { reportedUserId: alert.userId },
        take: 10,
        orderBy: { createdAt: 'desc' },
        include: {
          reporter: {
            select: {
              freelancerProfile: { select: { displayName: true } },
              clientProfile: { select: { displayName: true } },
            },
          },
        },
      }),
      db.report.count({
        where: { reporterId: alert.userId },
      }),
    ])

    // Count all fraud alerts for this user
    const totalFraudAlerts = await db.fraudAlert.count({
      where: { userId: alert.userId },
    })

    return apiSuccess({
      alert,
      userReportHistory: {
        reportsReceived,
        reportsFiledCount: reportsFiled,
      },
      totalFraudAlerts,
    })
  } catch (error: any) {
    console.error('Fraud alert detail error:', error)
    return apiError('Failed to fetch fraud alert', 500)
  }
}
