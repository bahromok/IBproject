import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// Severity priority map for ordering (CRITICAL first)
const SEVERITY_ORDER: Record<string, number> = {
  CRITICAL: 0,
  HIGH: 1,
  MEDIUM: 2,
  LOW: 3,
}

// GET /api/fraud-alerts - List all fraud alerts (Admin only)
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const severity = searchParams.get('severity')
    const alertType = searchParams.get('alertType')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {}
    if (status) where.status = status
    if (severity) where.severity = severity
    if (alertType) where.alertType = alertType

    const [alerts, total] = await Promise.all([
      db.fraudAlert.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          user: {
            select: {
              id: true,
              email: true,
              role: true,
              isSuspended: true,
              createdAt: true,
              freelancerProfile: { select: { displayName: true } },
              clientProfile: { select: { displayName: true } },
            },
          },
        },
      }),
      db.fraudAlert.count({ where }),
    ])

    // Sort client-side by severity (CRITICAL first) then by createdAt desc
    const sortedAlerts = [...alerts].sort((a, b) => {
      const sevA = SEVERITY_ORDER[a.severity] ?? 4
      const sevB = SEVERITY_ORDER[b.severity] ?? 4
      if (sevA !== sevB) return sevA - sevB
      return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
    })

    // Count by severity and status for stats
    const [pendingCount, highCount, criticalCount] = await Promise.all([
      db.fraudAlert.count({ where: { status: 'PENDING' } }),
      db.fraudAlert.count({ where: { severity: 'HIGH', status: 'PENDING' } }),
      db.fraudAlert.count({ where: { severity: 'CRITICAL', status: 'PENDING' } }),
    ])

    return apiSuccess({
      alerts: sortedAlerts,
      stats: {
        pending: pendingCount,
        highPending: highCount,
        criticalPending: criticalCount,
      },
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Fraud alerts list error:', error)
    return apiError('Failed to fetch fraud alerts', 500)
  }
}
