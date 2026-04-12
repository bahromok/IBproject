'use client'

import { useState, useEffect } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Progress } from '@/components/ui/progress'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency, timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import {
  Briefcase, DollarSign, FileText, PlusCircle,
  Users, MessageSquare, ArrowRight, Eye, Bookmark, Send,
} from 'lucide-react'
import { EmailVerificationBanner } from '@/components/shared/email-verification-banner'

interface ApiContract {
  id: string
  title: string
  status: string
  totalAmount: number
  contractType: string
  createdAt: string
  updatedAt: string
  freelancer: { id: string; displayName: string; avatar: string | null }
  client: { id: string; displayName: string; avatar: string | null }
  milestones: Array<{ id: string; title: string; amount: number; status: string }>
}

interface ApiJob {
  id: string
  title: string
  status: string
  budgetMin: number | null
  budgetMax: number | null
  budgetType: string
  proposalsCount: number
  createdAt: string
  updatedAt: string
  category: string
  client: { id: string; displayName: string; avatar: string | null; rating: number }
}

interface ApiTransaction {
  id: string
  type: string
  amount: number
  fromUserId: string
  toUserId: string
  status: string
  createdAt: string
}

function StatusBadge({ status }: { status: string }) {
  const variants: Record<string, string> = {
    PENDING: 'bg-yellow-100 text-yellow-700 border-yellow-200',
    SHORTLISTED: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    ACCEPTED: 'bg-green-100 text-green-700 border-green-200',
    REJECTED: 'bg-red-100 text-red-700 border-red-200',
    WITHDRAWN: 'bg-slate-100 text-slate-500 border-slate-200',
    ACTIVE: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    COMPLETED: 'bg-green-100 text-green-700 border-green-200',
    PAUSED: 'bg-amber-100 text-amber-700 border-amber-200',
    CANCELLED: 'bg-slate-100 text-slate-500 border-slate-200',
    DISPUTED: 'bg-red-100 text-red-700 border-red-200',
    open: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    in_progress: 'bg-amber-100 text-amber-700 border-amber-200',
    completed: 'bg-green-100 text-green-700 border-green-200',
    draft: 'bg-slate-100 text-slate-500 border-slate-200',
    closed: 'bg-slate-100 text-slate-500 border-slate-200',
    DRAFT: 'bg-slate-100 text-slate-500 border-slate-200',
    OPEN: 'bg-emerald-100 text-emerald-700 border-emerald-200',
  }
  const label = status.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
  return (
    <Badge variant="outline" className={`text-xs font-medium ${variants[status] || 'bg-slate-100 text-slate-600'}`}>
      {label}
    </Badge>
  )
}

export function ClientDashboard() {
  const { user } = useAuthStore()
  const { navigate } = useNavigationStore()
  const [contracts, setContracts] = useState<ApiContract[]>([])
  const [jobs, setJobs] = useState<ApiJob[]>([])
  const [transactions, setTransactions] = useState<ApiTransaction[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!user?.id) return
    Promise.all([
      authFetch('/api/contracts?limit=50').then(r => r.json()).catch(() => ({})),
      authFetch('/api/jobs?clientId=' + user.id + '&limit=50&sortBy=createdAt&sortOrder=desc').then(r => r.json()).catch(() => ({})),
      authFetch('/api/transactions?limit=50').then(r => r.json()).catch(() => ({})),
    ]).then(([contractsData, jobsData, transactionsData]) => {
      setContracts(contractsData?.contracts || [])
      setJobs(jobsData?.jobs || [])
      setTransactions(transactionsData?.transactions || [])
      setLoading(false)
    })
  }, [user?.id])

  // Computed stats from real data
  const activeContracts = contracts.filter(c => c.status === 'ACTIVE').length
  const totalSpent = transactions
    .filter(t => t.type === 'ESCROW_DEPOSIT' && t.fromUserId === user?.id && t.status === 'COMPLETED')
    .reduce((sum, t) => sum + t.amount, 0)
  const postedJobs = jobs.length
  const pendingProposals = jobs
    .filter(j => j.status === 'open' || j.status === 'OPEN')
    .reduce((sum, j) => sum + (j.proposalsCount || 0), 0)

  const displayName = user?.displayName || user?.name || user?.email || 'there'

  const getContractProgress = (c: ApiContract) => {
    if (!c.milestones?.length) return 0
    const completed = c.milestones.filter(m => m.status === 'PAID' || m.status === 'APPROVED').length
    return Math.round((completed / c.milestones.length) * 100)
  }

  const activeContractList = contracts.filter(c => c.status === 'ACTIVE' || c.status === 'PAUSED')
  const recentJobs = [...jobs].sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())

  const budgetDisplay = (job: ApiJob) => {
    if (job.budgetType === 'HOURLY') return 'Hourly'
    if (job.budgetMin && job.budgetMax) return `${formatCurrency(job.budgetMin)} - ${formatCurrency(job.budgetMax)}`
    if (job.budgetMin) return `${formatCurrency(job.budgetMin)}+`
    return 'Budget not set'
  }

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-64" />
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-32 rounded-xl" />)}
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <Skeleton className="h-80 rounded-xl" />
          <Skeleton className="h-80 rounded-xl" />
        </div>
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Email Verification Banner */}
      <EmailVerificationBanner />

      {/* Welcome Banner */}
      <div className="bg-gradient-to-r from-emerald-500 to-emerald-600 rounded-2xl p-6 text-white">
        <h1 className="text-2xl font-bold">
          Welcome back, {displayName}!
        </h1>
        <p className="text-emerald-100 mt-1">
          Here&apos;s what&apos;s happening with your projects.
        </p>
      </div>

      {/* Stats Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard
          label="Active Contracts"
          value={activeContracts}
          icon={Briefcase}
          iconColor="bg-emerald-100 text-emerald-600"
          sub="In progress"
        />
        <StatCard
          label="Total Spent"
          value={totalSpent}
          format="currency"
          icon={DollarSign}
          iconColor="bg-amber-100 text-amber-600"
          sub="All time"
        />
        <StatCard
          label="Posted Jobs"
          value={postedJobs}
          icon={FileText}
          iconColor="bg-violet-100 text-violet-600"
          sub={`${pendingProposals} pending proposals`}
        />
        <StatCard
          label="Pending Proposals"
          value={pendingProposals}
          icon={Send}
          iconColor="bg-rose-100 text-rose-600"
          sub="Awaiting review"
        />
      </div>

      {/* Quick Actions */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        {[
          { label: 'Post a Job', view: 'jobs/post' as const, icon: PlusCircle, color: 'bg-emerald-100 text-emerald-600' },
          { label: 'Find Freelancers', view: 'freelancers' as const, icon: Users, color: 'bg-amber-100 text-amber-600' },
          { label: 'Messages', view: 'messages' as const, icon: MessageSquare, color: 'bg-violet-100 text-violet-600' },
          { label: 'Saved Jobs', view: 'saved-jobs' as const, icon: Bookmark, color: 'bg-rose-100 text-rose-600' },
        ].map(action => (
          <Button
            key={action.label}
            variant="outline"
            onClick={() => navigate(action.view)}
            className="h-auto py-4 flex-col gap-2 border-slate-200 hover:border-emerald-200 hover:bg-emerald-50/50 transition-all group"
          >
            <div className={`w-10 h-10 rounded-lg flex items-center justify-center shrink-0 ${action.color}`}>
              <action.icon className="w-5 h-5" />
            </div>
            <span className="text-sm font-medium text-slate-700 group-hover:text-emerald-700">{action.label}</span>
          </Button>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Active Contracts */}
        <Card className="border-slate-200">
          <CardHeader className="pb-3 flex flex-row items-center justify-between">
            <CardTitle className="text-lg font-semibold text-slate-900">Active Contracts</CardTitle>
            <Button variant="ghost" size="sm" onClick={() => navigate('contracts')} className="text-emerald-600 hover:text-emerald-700 text-xs">
              View All <ArrowRight className="w-3 h-3 ml-1" />
            </Button>
          </CardHeader>
          <CardContent className="p-0">
            {activeContractList.length === 0 ? (
              <div className="px-6 pb-6">
                <EmptyState
                  icon={Briefcase}
                  title="No active contracts"
                  description="When you hire a freelancer, their contracts will appear here."
                  actionLabel="Find Freelancers"
                  onAction={() => navigate('freelancers')}
                />
              </div>
            ) : (
              <div className="max-h-96 overflow-y-auto">
                {activeContractList.slice(0, 5).map(contract => (
                  <div key={contract.id} className="flex items-center gap-3 p-4 hover:bg-slate-50 border-b border-slate-100 last:border-0 transition-colors">
                    <Avatar className="h-10 w-10 border border-slate-200 shrink-0">
                      <AvatarFallback className="bg-emerald-100 text-emerald-700 text-xs font-semibold">
                        {contract.freelancer?.displayName?.[0] || 'F'}
                      </AvatarFallback>
                    </Avatar>
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-slate-900 truncate">{contract.title}</div>
                      <div className="text-xs text-slate-500 mt-0.5">
                        {contract.freelancer?.displayName || 'Freelancer'} &middot; {formatCurrency(contract.totalAmount)}
                      </div>
                      {contract.milestones?.length > 0 && (
                        <div className="flex items-center gap-2 mt-2">
                          <Progress value={getContractProgress(contract)} className="h-1.5 flex-1" />
                          <span className="text-xs text-slate-400">
                            {contract.milestones.filter(m => m.status === 'PAID' || m.status === 'APPROVED').length}/{contract.milestones.length}
                          </span>
                        </div>
                      )}
                    </div>
                    <Button variant="ghost" size="sm" onClick={() => navigate('contracts/detail', { id: contract.id })} className="shrink-0 text-emerald-600 hover:text-emerald-700">
                      <Eye className="w-4 h-4" />
                    </Button>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>

        {/* Recent Jobs */}
        <Card className="border-slate-200">
          <CardHeader className="pb-3 flex flex-row items-center justify-between">
            <CardTitle className="text-lg font-semibold text-slate-900">Your Recent Jobs</CardTitle>
            <Button variant="ghost" size="sm" onClick={() => navigate('jobs/my-jobs')} className="text-emerald-600 hover:text-emerald-700 text-xs">
              View All <ArrowRight className="w-3 h-3 ml-1" />
            </Button>
          </CardHeader>
          <CardContent className="p-0">
            {recentJobs.length === 0 ? (
              <div className="px-6 pb-6">
                <EmptyState
                  icon={FileText}
                  title="No jobs posted"
                  description="Post your first job to start receiving proposals from talented freelancers."
                  actionLabel="Post a Job"
                  onAction={() => navigate('jobs/post')}
                />
              </div>
            ) : (
              <div className="max-h-96 overflow-y-auto">
                {recentJobs.slice(0, 5).map(job => (
                  <div
                    key={job.id}
                    className="flex items-center justify-between p-4 hover:bg-slate-50 border-b border-slate-100 last:border-0 transition-colors cursor-pointer"
                    onClick={() => navigate('jobs/detail', { id: job.id })}
                  >
                    <div className="flex-1 min-w-0">
                      <div className="text-sm font-medium text-slate-900 truncate">{job.title}</div>
                      <div className="flex items-center gap-3 mt-1">
                        <span className="text-xs text-emerald-600 font-medium">
                          {budgetDisplay(job)}
                        </span>
                        <span className="text-xs text-slate-400">
                          {timeAgo(job.createdAt)}
                        </span>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 ml-4">
                      <Badge variant="secondary" className="text-xs bg-slate-100 text-slate-600">
                        {job.proposalsCount} proposals
                      </Badge>
                      <StatusBadge status={job.status} />
                    </div>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>
      </div>
    </motion.div>
  )
}
