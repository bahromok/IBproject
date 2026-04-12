'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency } from '@/lib/format'
import { cn } from '@/lib/utils'
import { authFetch } from '@/lib/api-fetch'
import {
  Users,
  Briefcase,
  DollarSign,
  AlertTriangle,
  Scale,
  ChevronRight,
  UserPlus,
  Star,
  BarChart3,
  TrendingUp,
  ShieldAlert,
} from 'lucide-react'

interface AnalyticsData {
  users: { total: number; clients: number; freelancers: number; admins: number }
  jobs: { total: number; open: number; completed: number; inProgress: number }
  contracts: { total: number; active: number; completed: number; disputed: number }
  revenue: { totalEscrowed: number; estimatedPlatformFees: number }
  reviews: { total: number }
  reports: { total: number; pending: number }
  fraudAlerts: { total: number; high: number; critical: number }
  topCategories: { category: string; count: number }[]
  recentUsers: Array<{
    id: string
    email?: string
    role: string
    createdAt: string
    clientProfile?: { displayName: string } | null
    freelancerProfile?: { displayName: string } | null
  }>
}

const categoryLabels: Record<string, string> = {
  DESIGNER: 'Designer',
  WEB_DEVELOPER: 'Web Dev',
  VIDEO_EDITOR: 'Video Editor',
  CONTENT_CREATOR: 'Content Creator',
  OTHER: 'Other',
}

export function AdminDashboard() {
  const { navigate } = useNavigationStore()
  const [data, setData] = useState<AnalyticsData | null>(null)
  const [loading, setLoading] = useState(true)

  const fetchAnalytics = useCallback(async () => {
    setLoading(true)
    try {
      const res = await authFetch('/api/admin/analytics')
      if (res.ok) {
        const result = await res.json()
        setData(result)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { fetchAnalytics() }, [fetchAnalytics])

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-48" />
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {[1, 2, 3, 4].map(i => (
            <Skeleton key={i} className="h-28 rounded-xl" />
          ))}
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => (
            <Skeleton key={i} className="h-24 rounded-xl" />
          ))}
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <Skeleton className="h-64 rounded-xl" />
          <Skeleton className="h-64 rounded-xl" />
        </div>
      </div>
    )
  }

  if (!data) {
    return (
      <div className="space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Admin Dashboard</h1>
          <p className="text-slate-500 mt-1">Platform overview and management</p>
        </div>
        <EmptyState
          icon={AlertTriangle}
          title="Unable to load analytics"
          description="There was a problem fetching platform data. Please try again."
          actionLabel="Retry"
          onAction={fetchAnalytics}
        />
      </div>
    )
  }

  const totalJobs = data.jobs.total
  const activeContracts = data.contracts.active
  const completedContracts = data.contracts.completed
  const disputedContracts = data.contracts.disputed
  const pendingFraud = data.fraudAlerts?.total || 0

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-6"
    >
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Admin Dashboard</h1>
        <p className="text-slate-500 mt-1">Platform overview and management</p>
      </div>

      {/* Primary Stats Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard
          label="Total Users"
          value={data.users.total}
          icon={Users}
          iconColor="bg-emerald-100 text-emerald-600"
          sub={`${data.users.freelancers} freelancers, ${data.users.clients} clients`}
        />
        <StatCard
          label="Active Jobs"
          value={data.jobs.open}
          icon={Briefcase}
          iconColor="bg-amber-100 text-amber-600"
          sub={`${totalJobs} total jobs`}
        />
        <StatCard
          label="Pending Reports"
          value={data.reports.pending}
          icon={AlertTriangle}
          iconColor={data.reports.pending > 0 ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-600'}
          sub={`${data.reports.total} total reports`}
        />
        <StatCard
          label="Platform Revenue"
          value={data.revenue.estimatedPlatformFees}
          format="currency"
          icon={DollarSign}
          iconColor="bg-rose-100 text-rose-600"
          sub={`${formatCurrency(data.revenue.totalEscrowed)} total GMV`}
        />
      </div>

      {/* Secondary Stats */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Active Contracts"
          value={activeContracts}
          icon={Briefcase}
          iconColor="bg-violet-100 text-violet-600"
          sub={`${completedContracts} completed`}
        />
        <StatCard
          label="Open Disputes"
          value={disputedContracts}
          icon={Scale}
          iconColor={disputedContracts > 0 ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-600'}
        />
        <StatCard
          label="Reviews"
          value={data.reviews.total}
          icon={Star}
          iconColor="bg-amber-100 text-amber-600"
        />
        <StatCard
          label="Admins"
          value={data.users.admins}
          icon={UserPlus}
          iconColor="bg-cyan-100 text-cyan-600"
          sub={`${data.users.clients} clients`}
        />
      </div>

      {/* Recent Activity & Quick Actions */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Recent Users */}
        <Card>
          <CardHeader className="pb-2">
            <div className="flex items-center justify-between">
              <CardTitle className="text-base font-semibold">Recent Registrations</CardTitle>
              <Button variant="ghost" size="sm" onClick={() => navigate('admin/users')}>
                View All <ChevronRight className="w-3 h-3 ml-1" />
              </Button>
            </div>
          </CardHeader>
          <CardContent>
            {(data.recentUsers || []).length > 0 ? (
              <div className="space-y-3">
                {data.recentUsers.map((u) => (
                  <div key={u.id} className="flex items-center gap-3 p-2.5 bg-slate-50 rounded-lg">
                    <div className="w-8 h-8 rounded-full bg-emerald-100 flex items-center justify-center">
                      <span className="text-xs font-semibold text-emerald-700">
                        {((u.clientProfile?.displayName || u.freelancerProfile?.displayName || u.email || 'U')).slice(0, 2).toUpperCase()}
                      </span>
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-slate-900 truncate">
                        {u.clientProfile?.displayName || u.freelancerProfile?.displayName || u.email}
                      </div>
                      <div className="text-xs text-slate-500">{u.role} · {new Date(u.createdAt).toLocaleDateString()}</div>
                    </div>
                    <span className={cn(
                      'px-2 py-0.5 rounded-full text-[10px] font-medium',
                      u.role === 'CLIENT' ? 'bg-amber-100 text-amber-700' : u.role === 'ADMIN' ? 'bg-slate-100 text-slate-600' : 'bg-emerald-100 text-emerald-700'
                    )}>
                      {u.role}
                    </span>
                  </div>
                ))}
              </div>
            ) : (
              <EmptyState
                icon={Users}
                title="No users yet"
                description="New user registrations will appear here."
              />
            )}
          </CardContent>
        </Card>

        {/* Quick Actions */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Quick Actions</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 gap-3">
              {([
                { label: 'Manage Users', view: 'admin/users' as const, icon: Users, color: 'bg-emerald-100 text-emerald-600' },
                { label: 'Review Jobs', view: 'admin/jobs' as const, icon: Briefcase, color: 'bg-amber-100 text-amber-600' },
                { label: 'Pending Reports', view: 'admin/reports' as const, icon: AlertTriangle, color: 'bg-red-100 text-red-600', badge: data.reports.pending || 0 },
                { label: 'Fraud Alerts', view: 'admin/fraud' as const, icon: ShieldAlert, color: pendingFraud > 0 ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-600', badge: pendingFraud || 0 },
                { label: 'Open Disputes', view: 'admin/disputes' as const, icon: Scale, color: 'bg-violet-100 text-violet-600', badge: data.contracts.disputed || 0 },
                { label: 'Analytics', view: 'admin/analytics' as const, icon: BarChart3, color: 'bg-cyan-100 text-cyan-600' },
                { label: 'Platform Stats', view: 'admin/analytics' as const, icon: TrendingUp, color: 'bg-rose-100 text-rose-600' },
              ] as const).map((action) => (
                <button
                  key={action.label}
                  onClick={() => navigate(action.view)}
                  className="flex items-center gap-3 p-3 bg-slate-50 rounded-lg hover:bg-emerald-50 hover:border-emerald-200 border border-transparent transition-all text-left group"
                >
                  <div className={cn('w-9 h-9 rounded-lg flex items-center justify-center shrink-0', action.color)}>
                    <action.icon className="w-4 h-4" />
                  </div>
                  <div className="flex-1 min-w-0 flex items-center gap-2">
                    <span className="text-sm font-medium text-slate-700 group-hover:text-emerald-700">{action.label}</span>
                    {'badge' in action && action.badge > 0 && (
                      <div className="w-5 h-5 rounded-full bg-red-500 text-white text-[10px] font-bold flex items-center justify-center ml-auto">
                        {action.badge}
                      </div>
                    )}
                  </div>
                </button>
              ))}
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Top Categories */}
      {(data.topCategories || []).length > 0 && (
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Job Categories</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              {(data.topCategories || []).map(c => (
                <div key={c.category} className="p-3 bg-slate-50 rounded-lg text-center">
                  <div className="text-lg font-bold text-slate-900">{c.count}</div>
                  <div className="text-xs text-slate-500 mt-0.5">{categoryLabels[c.category] || c.category}</div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}
    </motion.div>
  )
}
