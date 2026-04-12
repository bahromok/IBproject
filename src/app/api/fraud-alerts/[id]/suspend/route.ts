import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// POST /api/fraud-alerts/[id]/suspend - Suspend user associated with fraud alert (Admin only)
export async function POST(
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
    const { adminNote } = body
    const reason = adminNote || `Suspended due to ${alert.alertType} fraud alert`

    // Suspend the user
    await db.user.update({
      where: { id: alert.userId },
      data: {
        isSuspended: true,
        suspendReason: reason,
      },
    })

    // Update fraud alert status
    const updatedAlert = await db.fraudAlert.update({
      where: { id },
      data: {
        status: 'ACTION_TAKEN',
        adminNote: reason,
        reviewedBy: user.id,
        reviewedAt: new Date(),
      },
    })

    return apiSuccess({
      alert: updatedAlert,
      message: `User has been suspended successfully`,
    })
  } catch (error: any) {
    console.error('Fraud alert suspend error:', error)
    return apiError('Failed to suspend user', 500)
  }
}
