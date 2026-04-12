import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { notifyReportStatus } from '@/lib/notify'

// GET /api/admin/reports - List all reports
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const { searchParams } = new URL(request.url)
    const status = searchParams.get('status')
    const relatedEntityType = searchParams.get('relatedEntityType')
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where: any = {}
    if (status) where.status = status
    if (relatedEntityType) where.relatedEntityType = relatedEntityType

    const [reports, total] = await Promise.all([
      db.report.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
        include: {
          reporter: {
            select: {
              id: true,
              email: true,
              clientProfile: { select: { displayName: true } },
              freelancerProfile: { select: { displayName: true } },
            },
          },
        },
      }),
      db.report.count({ where }),
    ])

    return apiSuccess({
      reports,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Admin reports list error:', error)
    return apiError('Failed to fetch reports', 500)
  }
}

// PUT /api/admin/reports - Update report status, add admin note
export async function PUT(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const body = await request.json()
    const { reportId, status, adminNote } = body

    if (!reportId) return apiError('Report ID is required')

    const report = await db.report.findUnique({ where: { id: reportId } })
    if (!report) return apiError('Report not found', 404)

    const validStatuses = ['PENDING', 'UNDER_REVIEW', 'RESOLVED', 'DISMISSED']
    if (status && !validStatuses.includes(status)) {
      return apiError(`Invalid status. Use: ${validStatuses.join(', ')}`)
    }

    const updated = await db.report.update({
      where: { id: reportId },
      data: {
        ...(status && { status }),
        ...(adminNote !== undefined && { adminNote }),
        ...(status === 'RESOLVED' || status === 'DISMISSED' ? { resolvedBy: user.id, resolvedAt: new Date() } : {}),
      },
    })

    // Notify reporter when report is resolved or dismissed (fire-and-forget)
    if (status === 'RESOLVED' || status === 'DISMISSED') {
      notifyReportStatus(report.reporterId, reportId, status)
    }

    return apiSuccess({ report: updated })
  } catch (error: any) {
    console.error('Admin report update error:', error)
    return apiError('Failed to update report', 500)
  }
}
