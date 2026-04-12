'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Skeleton } from '@/components/ui/skeleton'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table'
import {
  DollarSign,
  TrendingUp,
  Users,
  Briefcase,
  Download,
  Star,
  Shield,
  Clock,
  Heart,
  FileText,
  Activity,
  Scale,
  CheckCircle2,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { StatCard } from '@/components/shared/stat-card'
import { authFetch } from '@/lib/api-fetch'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency } from '@/lib/format'

interface AnalyticsFull {
  users: { total: number; clients: number; freelancers: number; admins: number }
  jobs: { total: number; open: number; completed: number; inProgress: number }
  contracts: { total: number; active: number; completed: number; disputed: number }
  revenue: { totalEscrowed: number; estimatedPlatformFees: number }
  reviews: { total: number }
  reports: { total: number; pending: number }
  topCategories: { category: string; count: number }[]
  recentUsers: any[]
  recentJobs: any[]
}

const CATEGORY_LABELS: Record<string, string> = {
  DESIGNER: 'Designer',
  WEB_DEVELOPER: 'Web Developer',
  VIDEO_EDITOR: 'Video Editor',
  CONTENT_CREATOR: 'Content Creator',
  OTHER: 'Other',
}

export function AdminAnalytics() {
  const [data, setData] = useState<AnalyticsFull | null>(null)
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

  const exportCSV = () => {
    if (!data) return
    const rows = [
      ['Metric', 'Value'],
      ['Total Users', data.users.total],
      ['Total Clients', data.users.clients],
      ['Total Freelancers', data.users.freelancers],
      ['Total Jobs', data.jobs.total],
      ['Open Jobs', data.jobs.open],
      ['Completed Jobs', data.jobs.completed],
      ['Active Contracts', data.contracts.active],
      ['Completed Contracts', data.contracts.completed],
      ['Disputed Contracts', data.contracts.disputed],
      ['Total Escrowed', data.revenue.totalEscrowed],
      ['Platform Fees', data.revenue.estimatedPlatformFees],
      ['Total Reviews', data.reviews.total],
      ['Pending Reports', data.reports.pending],
    ]
    const csv = rows.map(r => r.join(',')).join('\n')
    const blob = new Blob([csv], { type: 'text/csv' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `studenthire-analytics-${new Date().toISOString().split('T')[0]}.csv`
    a.click()
    URL.revokeObjectURL(url)
  }

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-48" />
        <div className="grid grid-cols-2 lg:grid-cols-3 gap-4">
          {[1, 2, 3, 4, 5, 6].map(i => (
            <Skeleton key={i} className="h-24 rounded-xl" />
          ))}
        </div>
      </div>
    )
  }

  if (!data) {
    return (
      <div className="space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Platform Analytics</h1>
          <p className="text-sm text-slate-500 mt-0.5">Detailed insights into platform performance</p>
        </div>
        <EmptyState
          icon={Activity}
          title="Unable to load analytics"
          description="There was a problem fetching platform data."
          actionLabel="Retry"
          onAction={fetchAnalytics}
        />
      </div>
    )
  }

  const gmv = data.revenue.totalEscrowed || 0
  const fees = data.revenue.estimatedPlatformFees || 0
  const avgJobValue = data.jobs.total > 0 ? Math.round(gmv / data.jobs.total) : 0
  const avgContractValue = data.contracts.total > 0 ? Math.round(gmv / data.contracts.total) : 0
  const disputeRate = data.contracts.total > 0 ? ((data.contracts.disputed / data.contracts.total) * 100).toFixed(1) : '0'
  const reportRate = data.users.total > 0 ? ((data.reports.total / data.users.total) * 100).toFixed(1) : '0'

  const categoryData = (data.topCategories || []).filter(c => c.count > 0)
  const completionRate = data.jobs.total > 0 ? Math.round((data.jobs.completed / data.jobs.total) * 100) : 0

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-6"
    >
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Platform Analytics</h1>
          <p className="text-sm text-slate-500 mt-0.5">Detailed insights into platform performance</p>
        </div>
        <Button variant="outline" size="sm" className="gap-1.5" onClick={exportCSV}>
          <Download className="w-3.5 h-3.5" /> Export CSV
        </Button>
      </div>

      {/* Key Metrics */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        <StatCard
          label="Platform Revenue (Fees)"
          value={fees}
          format="currency"
          icon={DollarSign}
          iconColor="bg-emerald-100 text-emerald-600"
          sub={`${formatCurrency(gmv)} total GMV`}
        />
        <StatCard
          label="Avg Job Value"
          value={avgJobValue}
          format="currency"
          icon={Briefcase}
          iconColor="bg-amber-100 text-amber-600"
          sub={`${data.jobs.total} total jobs`}
        />
        <StatCard
          label="Avg Contract Value"
          value={avgContractValue}
          format="currency"
          icon={FileText}
          iconColor="bg-violet-100 text-violet-600"
          sub={`${data.contracts.total} total contracts`}
        />
        <StatCard
          label="Clients"
          value={data.users.clients}
          icon={Users}
          iconColor="bg-rose-100 text-rose-600"
          sub={`${data.users.freelancers} freelancers`}
        />
        <StatCard
          label="Completion Rate"
          value={completionRate}
          suffix="%"
          icon={CheckCircle2}
          iconColor="bg-emerald-100 text-emerald-600"
          sub={`${data.jobs.completed} completed`}
        />
        <StatCard
          label="Total Reviews"
          value={data.reviews.total}
          icon={Star}
          iconColor="bg-amber-100 text-amber-600"
        />
      </div>

      {/* Job Categories */}
      {categoryData.length > 0 && (
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Jobs by Category</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-3">
              {categoryData.map(c => (
                <div key={c.category} className="p-4 bg-slate-50 rounded-lg text-center">
                  <div className="text-xl font-bold text-slate-900">{c.count}</div>
                  <div className="text-xs text-slate-500 mt-1">{CATEGORY_LABELS[c.category] || c.category}</div>
                  <div className="w-full bg-slate-200 rounded-full h-1.5 mt-2">
                    <div
                      className="bg-emerald-500 h-1.5 rounded-full"
                      style={{ width: `${data.jobs.total > 0 ? (c.count / data.jobs.total) * 100 : 0}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      {/* Platform Health + Recent Activity */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Platform Health */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold flex items-center gap-2">
              <Activity className="w-4 h-4 text-emerald-500" />
              Platform Health
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-4">
              {[
                {
                  label: 'Dispute Rate',
                  value: `${disputeRate}%`,
                  healthy: parseFloat(disputeRate) < 10,
                  icon: Scale,
                },
                {
                  label: 'Active Contracts',
                  value: `${data.contracts.active}`,
                  healthy: data.contracts.active > 0,
                  icon: FileText,
                },
                {
                  label: 'Job Completion Rate',
                  value: `${completionRate}%`,
                  healthy: completionRate > 50,
                  icon: CheckCircle2,
                },
                {
                  label: 'Report Rate',
                  value: `${reportRate}%`,
                  healthy: parseFloat(reportRate) < 10,
                  icon: Shield,
                },
              ].map((metric) => (
                <div key={metric.label} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                  <div className="flex items-center gap-3">
                    <metric.icon className="w-4 h-4 text-slate-500" />
                    <span className="text-sm text-slate-700">{metric.label}</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-semibold text-slate-900">{metric.value}</span>
                    <span className={cn(
                      'px-2 py-0.5 rounded-full text-[10px] font-medium',
                      metric.healthy
                        ? 'bg-emerald-100 text-emerald-700'
                        : 'bg-amber-100 text-amber-700'
                    )}>
                      {metric.healthy ? 'Healthy' : 'Monitor'}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>

        {/* Contract Breakdown */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Contract Breakdown</CardTitle>
          </CardHeader>
          <CardContent>
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Status</TableHead>
                  <TableHead className="text-right">Count</TableHead>
                  <TableHead className="text-right">Percentage</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {[
                  { label: 'Active', count: data.contracts.active, color: 'bg-emerald-500' },
                  { label: 'Completed', count: data.contracts.completed, color: 'bg-teal-500' },
                  { label: 'Disputed', count: data.contracts.disputed, color: 'bg-red-500' },
                ].map((row) => (
                  <TableRow key={row.label}>
                    <TableCell>
                      <div className="flex items-center gap-2">
                        <div className={cn('w-2.5 h-2.5 rounded-full', row.color)} />
                        <span className="text-sm font-medium text-slate-900">{row.label}</span>
                      </div>
                    </TableCell>
                    <TableCell className="text-right text-sm font-semibold text-slate-900">{row.count}</TableCell>
                    <TableCell className="text-right text-sm text-slate-500">
                      {data.contracts.total > 0
                        ? `${((row.count / data.contracts.total) * 100).toFixed(1)}%`
                        : '0%'}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>

            {/* Visual bar */}
            <div className="mt-4 flex h-3 rounded-full overflow-hidden">
              {data.contracts.total > 0 && (
                <>
                  <div
                    className="bg-emerald-500"
                    style={{ width: `${(data.contracts.active / data.contracts.total) * 100}%` }}
                  />
                  <div
                    className="bg-teal-500"
                    style={{ width: `${(data.contracts.completed / data.contracts.total) * 100}%` }}
                  />
                  <div
                    className="bg-red-500"
                    style={{ width: `${(data.contracts.disputed / data.contracts.total) * 100}%` }}
                  />
                </>
              )}
              {data.contracts.total === 0 && (
                <div className="bg-slate-200 w-full" />
              )}
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Recent Users */}
      {(data.recentUsers || []).length > 0 && (
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Recent Registrations</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>User</TableHead>
                    <TableHead>Role</TableHead>
                    <TableHead className="text-right">Joined</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {data.recentUsers.map((u) => (
                    <TableRow key={u.id}>
                      <TableCell>
                        <div className="flex items-center gap-3">
                          <div className={cn(
                            'w-8 h-8 rounded-full flex items-center justify-center text-xs font-semibold',
                            u.role === 'CLIENT' ? 'bg-amber-100 text-amber-700' : u.role === 'ADMIN' ? 'bg-slate-200 text-slate-700' : 'bg-emerald-100 text-emerald-700'
                          )}>
                            {((u.clientProfile?.displayName || u.freelancerProfile?.displayName || u.email || 'U')).slice(0, 2).toUpperCase()}
                          </div>
                          <span className="text-sm font-medium text-slate-900">
                            {u.clientProfile?.displayName || u.freelancerProfile?.displayName || u.email}
                          </span>
                        </div>
                      </TableCell>
                      <TableCell>
                        <span className={cn(
                          'px-2 py-0.5 rounded-full text-xs font-medium',
                          u.role === 'CLIENT' ? 'bg-amber-100 text-amber-700' : u.role === 'ADMIN' ? 'bg-slate-100 text-slate-600' : 'bg-emerald-100 text-emerald-700'
                        )}>
                          {u.role}
                        </span>
                      </TableCell>
                      <TableCell className="text-right text-sm text-slate-500">
                        {new Date(u.createdAt).toLocaleDateString()}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Recent Jobs */}
      {(data.recentJobs || []).length > 0 && (
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base font-semibold">Recent Jobs</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Title</TableHead>
                    <TableHead>Category</TableHead>
                    <TableHead>Status</TableHead>
                    <TableHead className="text-right">Created</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {data.recentJobs.map((job) => (
                    <TableRow key={job.id}>
                      <TableCell className="text-sm font-medium text-slate-900">{job.title}</TableCell>
                      <TableCell className="text-sm text-slate-500">
                        {CATEGORY_LABELS[job.category] || job.category}
                      </TableCell>
                      <TableCell>
                        <span className={cn(
                          'px-2 py-0.5 rounded-full text-xs font-medium',
                          job.status === 'open' || job.status === 'OPEN'
                            ? 'bg-emerald-100 text-emerald-700'
                            : 'bg-slate-100 text-slate-600'
                        )}>
                          {job.status}
                        </span>
                      </TableCell>
                      <TableCell className="text-right text-sm text-slate-500">
                        {new Date(job.createdAt).toLocaleDateString()}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      )}
    </motion.div>
  )
}
