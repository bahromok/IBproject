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
  Briefcase, DollarSign, FileText, TrendingUp,
  Search, ArrowRight, Eye, User, Rocket,
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
  client: { id: string; displayName: string; avatar: string | null }
  freelancer: { id: string; displayName: string; avatar: string | null }
  milestones: Array<{ id: string; title: string; amount: number; status: string }>
}

interface ApiJob {
  id: string
  title: string
  budgetMin: number | null
  budgetMax: number | null
  budgetType: string
  category: string
  status: string
  proposalsCount: number
  createdAt: string
  client: { id: string; displayName: string; avatar: string | null }
  skills: string[]
}

interface ApiTransaction {
  id: string
  type: string
  amount: number
  netAmount: number
  fromUserId: string
  toUserId: string
  status: string
  createdAt: string
}

interface ApiProposal {
  id: string
  status: string
  proposedAmount: number
  createdAt: string
}

export function FreelancerDashboard() {
  const { user } = useAuthStore()
  const { navigate } = useNavigationStore()
  const [contracts, setContracts] = useState<ApiContract[]>([])
  const [proposals, setProposals] = useState<ApiProposal[]>([])
  const [recommendedJobs, setRecommendedJobs] = useState<ApiJob[]>([])
  const [transactions, setTransactions] = useState<ApiTransaction[]>([])
  const [loading, setLoading] = useState(true)

  const displayName = user?.displayName || user?.name || user?.email || 'there'

  useEffect(() => {
    if (!user?.id) return
    Promise.all([
      authFetch('/api/contracts?limit=50').then(r => r.json()).catch(() => ({})),
      authFetch('/api/proposals?limit=50').then(r => r.json()).catch(() => ({})),
      authFetch('/api/jobs?limit=8&sortBy=createdAt&sortOrder=desc').then(r => r.json()).catch(() => ({})),
      authFetch('/api/transactions?limit=50').then(r => r.json()).catch(() => ({})),
    ]).then(([contractsData, proposalsData, jobsData, transactionsData]) => {
      setContracts(contractsData?.contracts || [])
      setProposals(proposalsData?.proposals || [])
      setRecommendedJobs(jobsData?.jobs || [])
      setTransactions(transactionsData?.transactions || [])
      setLoading(false)
    })
  }, [user?.id])

  // Computed stats from real data
  const activeContracts = contracts.filter(c => c.status === 'ACTIVE').length
  const totalEarnings = transactions
    .filter(t => t.type === 'PAYMENT_RELEASE' && t.toUserId === user?.id && t.status === 'COMPLETED')
    .reduce((sum, t) => sum + (t.netAmount || t.amount), 0)
  const profileViews = 0 // No tracking implemented
  const completedContracts = contracts.filter(c => c.status === 'COMPLETED').length
  const successRate = contracts.length > 0
    ? Math.round((completedContracts / contracts.length) * 100)
    : 0

  const budgetDisplay = (job: ApiJob) => {
    if (job.budgetType === 'HOURLY') return 'Hourly'
    if (job.budgetMin && job.budgetMax) return `${formatCurrency(job.budgetMin)} - ${formatCurrency(job.budgetMax)}`
    if (job.budgetMin) return `${formatCurrency(job.budgetMin)}+`
    return 'Budget not set'
  }

  const activeContractList = contracts.filter(c => c.status === 'ACTIVE' || c.status === 'PAUSED')

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-28 w-full rounded-2xl" />
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
          Here&apos;s an overview of your freelance activity.
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
          label="Total Earnings"
          value={totalEarnings}
          format="currency"
          icon={DollarSign}
          iconColor="bg-amber-100 text-amber-600"
          sub="All time"
        />
        <StatCard
          label="Profile Views"
          value={profileViews}
          icon={Eye}
          iconColor="bg-violet-100 text-violet-600"
          sub="No tracking yet"
        />
        <StatCard
          label="Success Rate"
          value={successRate}
          suffix="%"
          icon={TrendingUp}
          iconColor="bg-rose-100 text-rose-600"
          sub={`${completedContracts}/${contracts.length} completed`}
        />
      </div>

      {/* Quick Actions */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        {[
          { label: 'Find Jobs', view: 'jobs' as const, icon: Search, color: 'bg-emerald-100 text-emerald-600' },
          { label: 'My Proposals', view: 'proposals' as const, icon: FileText, color: 'bg-amber-100 text-amber-600' },
          { label: 'My Contracts', view: 'contracts' as const, icon: Briefcase, color: 'bg-violet-100 text-violet-600' },
          { label: 'Profile', view: 'profile' as const, icon: User, color: 'bg-rose-100 text-rose-600' },
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

      {/* Earnings Section */}
      <Card className="border-slate-200">
        <CardHeader className="pb-3">
          <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
            <DollarSign className="w-5 h-5 text-emerald-500" />
            Earnings Overview
          </CardTitle>
        </CardHeader>
        <CardContent>
          {totalEarnings > 0 ? (
            <div className="flex items-center gap-6 p-4 bg-emerald-50 rounded-xl">
              <div className="w-14 h-14 rounded-2xl bg-emerald-100 flex items-center justify-center shrink-0">
                <DollarSign className="w-7 h-7 text-emerald-600" />
              </div>
              <div>
                <div className="text-3xl font-bold text-emerald-700">{formatCurrency(totalEarnings)}</div>
                <div className="text-sm text-emerald-600 mt-0.5">Total earnings to date</div>
              </div>
              <div className="ml-auto text-right hidden sm:block">
                <div className="text-sm text-slate-500">{activeContracts} active contract{activeContracts !== 1 ? 's' : ''}</div>
                <div className="text-sm text-slate-500">{completedContracts} completed</div>
              </div>
            </div>
          ) : (
            <div className="flex flex-col items-center justify-center py-10">
              <div className="w-16 h-16 rounded-2xl bg-emerald-50 flex items-center justify-center mb-4">
                <Rocket className="w-8 h-8 text-emerald-400" />
              </div>
              <h3 className="text-lg font-semibold text-slate-900 mb-1">Start earning!</h3>
              <p className="text-sm text-slate-500 text-center max-w-sm mb-4">
                Complete projects to build your earnings history and see your growth over time.
              </p>
              <Button onClick={() => navigate('jobs')} className="bg-emerald-500 hover:bg-emerald-600 text-white">
                <Search className="w-4 h-4 mr-2" />
                Browse Jobs
              </Button>
            </div>
          )}
        </CardContent>
      </Card>

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
                  description="Apply to jobs to start working on projects and earning money."
                  actionLabel="Browse Jobs"
                  onAction={() => navigate('jobs')}
                />
              </div>
            ) : (
              <div className="max-h-96 overflow-y-auto">
                {activeContractList.slice(0, 5).map(contract => {
                  const completed = contract.milestones?.filter(m => m.status === 'PAID' || m.status === 'APPROVED').length || 0
                  const total = contract.milestones?.length || 1
                  return (
                    <div
                      key={contract.id}
                      className="flex items-center gap-3 p-4 hover:bg-slate-50 border-b border-slate-100 last:border-0 transition-colors cursor-pointer"
                      onClick={() => navigate('contracts/detail', { id: contract.id })}
                    >
                      <Avatar className="h-10 w-10 border border-slate-200 shrink-0">
                        <AvatarFallback className="bg-violet-100 text-violet-700 text-xs font-semibold">
                          {contract.client?.displayName?.[0] || 'C'}
                        </AvatarFallback>
                      </Avatar>
                      <div className="flex-1 min-w-0">
                        <div className="text-sm font-medium text-slate-900 truncate">{contract.title}</div>
                        <div className="text-xs text-slate-500 mt-0.5">{contract.client?.displayName || 'Client'}</div>
                        <div className="flex items-center gap-2 mt-1.5">
                          <Progress value={(completed / total) * 100} className="h-1.5 flex-1" />
                          <span className="text-xs text-slate-400 shrink-0">{completed}/{total}</span>
                        </div>
                      </div>
                      <span className="text-sm font-semibold text-emerald-600 shrink-0">
                        {formatCurrency(contract.totalAmount)}
                      </span>
                    </div>
                  )
                })}
              </div>
            )}
          </CardContent>
        </Card>

        {/* Recommended Jobs */}
        <Card className="border-slate-200">
          <CardHeader className="pb-3 flex flex-row items-center justify-between">
            <CardTitle className="text-lg font-semibold text-slate-900">Recommended Jobs</CardTitle>
            <Button variant="ghost" size="sm" onClick={() => navigate('jobs')} className="text-emerald-600 hover:text-emerald-700 text-xs">
              Browse All <ArrowRight className="w-3 h-3 ml-1" />
            </Button>
          </CardHeader>
          <CardContent className="p-0">
            {recommendedJobs.length === 0 ? (
              <div className="px-6 pb-6">
                <EmptyState
                  icon={Search}
                  title="No jobs available"
                  description="Check back later for new opportunities matching your skills."
                />
              </div>
            ) : (
              <div className="max-h-96 overflow-y-auto">
                {recommendedJobs.slice(0, 4).map(job => (
                  <div key={job.id} className="p-4 hover:bg-slate-50 border-b border-slate-100 last:border-0 transition-colors">
                    <div
                      className="cursor-pointer"
                      onClick={() => navigate('jobs/detail', { id: job.id })}
                    >
                      <div className="text-sm font-medium text-slate-900 truncate hover:text-emerald-600">{job.title}</div>
                      <div className="flex items-center gap-2 mt-1">
                        <span className="text-xs text-emerald-600 font-medium">
                          {budgetDisplay(job)}
                        </span>
                        <span className="text-xs text-slate-300">&middot;</span>
                        <span className="text-xs text-slate-500">{job.client?.displayName || 'Client'}</span>
                        <span className="text-xs text-slate-300">&middot;</span>
                        <span className="text-xs text-slate-400">{timeAgo(job.createdAt)}</span>
                      </div>
                      {job.skills?.length > 0 && (
                        <div className="flex flex-wrap gap-1 mt-2">
                          {job.skills.slice(0, 3).map((skill: string) => (
                            <Badge key={skill} variant="secondary" className="text-xs bg-slate-100 text-slate-600 px-1.5 py-0">{skill}</Badge>
                          ))}
                          {job.skills.length > 3 && (
                            <Badge variant="secondary" className="text-xs bg-slate-100 text-slate-400 px-1.5 py-0">+{job.skills.length - 3}</Badge>
                          )}
                        </div>
                      )}
                    </div>
                    <div className="flex items-center justify-between mt-3">
                      <span className="text-xs text-slate-400">{job.proposalsCount} proposals</span>
                      <Button size="sm" variant="outline" onClick={() => navigate('jobs/detail', { id: job.id })} className="text-xs h-7 border-emerald-200 text-emerald-600 hover:bg-emerald-50">
                        Apply Now
                      </Button>
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
