import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'
import { checkAndFlagUser } from '@/lib/fraud-detection'

// POST /api/reports - File a report
export async function POST(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const body = await request.json()
    const { targetType, targetId, reportType, description } = body

    // Support both reportType and reason for backwards compatibility
    const resolvedReportType = reportType || body.reason
    const resolvedDescription = description || ''

    if (!targetType || !targetId || !resolvedReportType) {
      return apiError('Target type, target ID, and report type are required')
    }

    const validReportTypes = ['SCAM', 'HARASSMENT', 'SPAM', 'INAPPROPRIATE_CONTENT', 'PAYMENT_DISPUTE', 'OTHER']
    if (!validReportTypes.includes(resolvedReportType)) {
      return apiError(`Invalid report type. Use: ${validReportTypes.join(', ')}`)
    }

    if (!['USER', 'JOB', 'PROPOSAL', 'CONTRACT', 'MESSAGE'].includes(targetType.toUpperCase())) {
      return apiError('Target type must be USER, JOB, PROPOSAL, CONTRACT, or MESSAGE')
    }

    // Resolve reportedUserId from target
    let reportedUserId = targetId
    if (targetType.toUpperCase() === 'JOB') {
      const job = await db.job.findUnique({ where: { id: targetId } })
      if (!job) return apiError('Job not found', 404)
      reportedUserId = job.clientId
    } else if (targetType.toUpperCase() === 'PROPOSAL') {
      const proposal = await db.proposal.findUnique({
        where: { id: targetId },
      })
      if (!proposal) return apiError('Proposal not found', 404)
      reportedUserId = proposal.freelancerId
    } else if (targetType.toUpperCase() === 'USER') {
      if (targetId === user.id) return apiError('Cannot report yourself', 400)
      const targetUser = await db.user.findUnique({ where: { id: targetId } })
      if (!targetUser) return apiError('User not found', 404)
    } else if (targetType.toUpperCase() === 'CONTRACT') {
      const contract = await db.contract.findUnique({ where: { id: targetId } })
      if (!contract) return apiError('Contract not found', 404)
      // Report the other party in the contract
      reportedUserId = contract.clientId === user.id ? contract.freelancerId : contract.clientId
    } else if (targetType.toUpperCase() === 'MESSAGE') {
      const message = await db.message.findUnique({ where: { id: targetId } })
      if (!message) return apiError('Message not found', 404)
      reportedUserId = message.senderId === user.id ? 'unknown' : message.senderId
    }

    const report = await db.report.create({
      data: {
        reporterId: user.id,
        reportedUserId,
        reportType: resolvedReportType as any,
        description: resolvedDescription,
        relatedEntityType: targetType.toUpperCase() as any,
        relatedEntityId: targetId,
      },
    })

    // Fire-and-forget: check if reported user shows fraud patterns
    checkAndFlagUser(reportedUserId).catch(() => {})

    return apiSuccess({ report }, 201)
  } catch (error: any) {
    console.error('Report creation error:', error)
    return apiError('Failed to file report', 500)
  }
}

// GET /api/reports - Get user's reports
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult

    const { searchParams } = new URL(request.url)
    const page = parseInt(searchParams.get('page') || '1')
    const limit = parseInt(searchParams.get('limit') || '20')

    const where = { reporterId: user.id }

    const [reports, total] = await Promise.all([
      db.report.findMany({
        where,
        skip: (page - 1) * limit,
        take: limit,
        orderBy: { createdAt: 'desc' },
      }),
      db.report.count({ where }),
    ])

    return apiSuccess({
      reports,
      pagination: { page, limit, total, totalPages: Math.ceil(total / limit) },
    })
  } catch (error: any) {
    console.error('Reports fetch error:', error)
    return apiError('Failed to fetch reports', 500)
  }
}
