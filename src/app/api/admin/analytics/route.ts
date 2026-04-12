import { NextRequest } from 'next/server'
import { db } from '@/lib/db'
import { apiError, apiSuccess, getAuthenticatedUser, isAuthUser } from '@/lib/api-auth'

// GET /api/admin/analytics - Platform statistics
export async function GET(request: NextRequest) {
  try {
    const authResult = await getAuthenticatedUser(request)
    if (!isAuthUser(authResult)) return authResult
    const user = authResult
    if (user.role !== 'ADMIN') return apiError('Admin access required', 403)

    const [
      totalUsers,
      totalClients,
      totalFreelancers,
      totalJobs,
      activeJobs,
      completedJobs,
      totalContracts,
      activeContracts,
      completedContracts,
      disputedContracts,
      totalTransactions,
      platformRevenue,
      totalReviews,
      totalReports,
      pendingReports,
      pendingFraudAlerts,
      highFraudAlerts,
      criticalFraudAlerts,
      recentUsers,
      recentJobs,
    ] = await Promise.all([
      db.user.count(),
      db.user.count({ where: { role: 'CLIENT' } }),
      db.user.count({ where: { role: 'FREELANCER' } }),
      db.job.count({ where: { status: { not: 'DELETED' } } }),
      db.job.count({ where: { status: 'OPEN' } }),
      db.job.count({ where: { status: 'COMPLETED' } }),
      db.contract.count(),
      db.contract.count({ where: { status: 'ACTIVE' } }),
      db.contract.count({ where: { status: 'COMPLETED' } }),
      db.contract.count({ where: { status: 'DISPUTED' } }),
      db.transaction.count({ where: { status: 'COMPLETED' } }),
      db.transaction.aggregate({
        where: { type: 'ESCROW_DEPOSIT', status: 'COMPLETED' },
        _sum: { amount: true },
      }),
      db.review.count(),
      db.report.count(),
      db.report.count({ where: { status: 'PENDING' } }),
      db.fraudAlert.count({ where: { status: 'PENDING' } }),
      db.fraudAlert.count({ where: { severity: 'HIGH', status: 'PENDING' } }),
      db.fraudAlert.count({ where: { severity: 'CRITICAL', status: 'PENDING' } }),
      db.user.findMany({
        take: 5,
        orderBy: { createdAt: 'desc' },
        select: {
          id: true,
          email: true,
          role: true,
          createdAt: true,
          clientProfile: { select: { displayName: true } },
          freelancerProfile: { select: { displayName: true } },
        },
      }),
      db.job.findMany({
        take: 5,
        orderBy: { createdAt: 'desc' },
        select: {
          id: true,
          title: true,
          status: true,
          category: true,
          createdAt: true,
          proposalsCount: true,
        },
      }),
    ])

    // Calculate platform fee revenue (10% of service fees)
    const feeTransactions = await db.transaction.findMany({
      where: { type: 'ESCROW_DEPOSIT', status: 'COMPLETED' },
      select: { amount: true },
    })
    const totalEscrowed = feeTransactions.reduce((sum, t) => sum + t.amount, 0)
    const estimatedRevenue = totalEscrowed * 0.1

    // Top categories
    const jobCategories = await db.job.groupBy({
      by: ['category'],
      where: { status: { not: 'DELETED' } },
      _count: { category: true },
      orderBy: { _count: { category: 'desc' } },
      take: 10,
    })

    return apiSuccess({
      users: {
        total: totalUsers,
        clients: totalClients,
        freelancers: totalFreelancers,
        admins: totalUsers - totalClients - totalFreelancers,
      },
      jobs: {
        total: totalJobs,
        open: activeJobs,
        completed: completedJobs,
        inProgress: totalJobs - activeJobs - completedJobs,
      },
      contracts: {
        total: totalContracts,
        active: activeContracts,
        completed: completedContracts,
        disputed: disputedContracts,
      },
      revenue: {
        totalEscrowed: Math.round(totalEscrowed * 100) / 100,
        estimatedPlatformFees: Math.round(estimatedRevenue * 100) / 100,
      },
      reviews: { total: totalReviews },
      reports: {
        total: totalReports,
        pending: pendingReports,
      },
      fraudAlerts: {
        total: pendingFraudAlerts,
        high: highFraudAlerts,
        critical: criticalFraudAlerts,
      },
      topCategories: jobCategories.map((c) => ({
        category: c.category,
        count: c._count.category,
      })),
      recentUsers,
      recentJobs,
    })
  } catch (error: any) {
    console.error('Admin analytics error:', error)
    return apiError('Failed to fetch analytics', 500)
  }
}
