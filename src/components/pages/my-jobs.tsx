'use client'

import { useState, useEffect } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency, timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import {
  Briefcase, PlusCircle, Clock, FileText, Eye,
  XCircle
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'

interface Job {
  id: string
  title: string
  description: string
  status: string
  category: string
  budgetType: string
  budgetMin: number | null
  budgetMax: number | null
  experienceLevel: string
  proposalsCount: number
  createdAt: string
  updatedAt: string
}

function StatusBadge({ status }: { status: string }) {
  const variants: Record<string, string> = {
    open: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    OPEN: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    in_progress: 'bg-amber-100 text-amber-700 border-amber-200',
    completed: 'bg-green-100 text-green-700 border-green-200',
    draft: 'bg-slate-100 text-slate-500 border-slate-200',
    DRAFT: 'bg-slate-100 text-slate-500 border-slate-200',
    closed: 'bg-slate-100 text-slate-500 border-slate-200',
    cancelled: 'bg-red-100 text-red-700 border-red-200',
  }
  const label = status === 'in_progress' ? 'In Progress'
    : status === 'DRAFT' ? 'Draft'
    : status.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
  return (
    <Badge variant="outline" className={`text-xs font-medium ${variants[status] || 'bg-slate-100 text-slate-600'}`}>
      {label}
    </Badge>
  )
}

export function MyJobs() {
  const { user } = useAuthStore()
  const { navigate } = useNavigationStore()
  const { toast } = useToast()
  const [jobs, setJobs] = useState<Job[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!user?.id) return
    authFetch('/api/jobs?clientId=' + user.id + '&limit=50&sortBy=createdAt&sortOrder=desc')
      .then(r => r.json())
      .then(data => {
        setJobs(data?.jobs || [])
        setLoading(false)
      })
      .catch(() => {
        setJobs([])
        setLoading(false)
      })
  }, [user?.id])

  const handleClose = async (jobId: string) => {
    try {
      const res = await authFetch(`/api/jobs/${jobId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'closed' }),
      })
      if (res.ok) {
        toast({ title: 'Job closed', description: 'The job has been closed successfully.' })
        setJobs(prev => prev.map(j => j.id === jobId ? { ...j, status: 'closed' } : j))
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to close job', variant: 'destructive' })
    }
  }

  const openJobs = jobs.filter(j => j.status === 'open' || j.status === 'OPEN')
  const inProgressJobs = jobs.filter(j => j.status === 'in_progress')
  const completedJobs = jobs.filter(j => j.status === 'completed')
  const draftJobs = jobs.filter(j => j.status === 'draft' || j.status === 'DRAFT')
  const totalProposals = jobs.reduce((sum, j) => sum + (j.proposalsCount || 0), 0)

  const filteredJobs = (tab: string) => {
    if (tab === 'open') return openJobs
    if (tab === 'in-progress') return inProgressJobs
    if (tab === 'completed') return completedJobs
    if (tab === 'draft') return draftJobs
    return jobs
  }

  const budgetDisplay = (job: Job) => {
    if (job.budgetType === 'HOURLY') return 'Hourly'
    if (job.budgetMin && job.budgetMax) return `${formatCurrency(job.budgetMin)} - ${formatCurrency(job.budgetMax)}`
    if (job.budgetMin) return `${formatCurrency(job.budgetMin)}+`
    return 'Budget not set'
  }

  const JobCard = ({ job }: { job: Job }) => (
    <Card className="border-slate-200 hover:shadow-sm transition-shadow">
      <CardContent className="p-4">
        <div className="flex items-start justify-between gap-4">
          <div className="flex-1 min-w-0">
            <div
              className="text-sm font-semibold text-slate-900 hover:text-emerald-600 cursor-pointer truncate"
              onClick={() => navigate('jobs/detail', { id: job.id })}
            >
              {job.title}
            </div>
            <div className="flex items-center gap-3 mt-1.5 flex-wrap">
              <span className="text-sm font-medium text-emerald-600">{budgetDisplay(job)}</span>
              <StatusBadge status={job.status} />
              <Badge variant="secondary" className="text-xs bg-slate-100 text-slate-600">
                <FileText className="w-3 h-3 mr-1" />
                {job.proposalsCount} proposals
              </Badge>
              <span className="text-xs text-slate-400 flex items-center gap-1">
                <Clock className="w-3 h-3" />
                {timeAgo(job.createdAt)}
              </span>
            </div>
          </div>
          <div className="flex items-center gap-2 shrink-0">
            <Button
              variant="outline"
              size="sm"
              onClick={() => navigate('jobs/detail', { id: job.id })}
              className="text-xs h-8 border-slate-200 text-slate-600 hover:bg-slate-50"
            >
              <Eye className="w-3 h-3 mr-1" />
              View
            </Button>
            {(job.status === 'open' || job.status === 'OPEN') && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => handleClose(job.id)}
                className="text-xs h-8 border-red-200 text-red-600 hover:bg-red-50"
              >
                <XCircle className="w-3 h-3 mr-1" />
                Close
              </Button>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  )

  if (loading) {
    return (
      <div className="space-y-6">
        <div className="flex items-center justify-between">
          <div>
            <Skeleton className="h-8 w-32 mb-2" />
            <Skeleton className="h-4 w-48" />
          </div>
          <Skeleton className="h-10 w-36" />
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-20 rounded-xl" />)}
        </div>
        <Skeleton className="h-10 w-64" />
        {[1, 2, 3].map(i => <Skeleton key={i} className="h-20 rounded-xl" />)}
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">My Jobs</h1>
          <p className="text-slate-500 mt-1">Manage all your posted job listings.</p>
        </div>
        <Button onClick={() => navigate('jobs/post')} className="bg-emerald-500 hover:bg-emerald-600 text-white shadow-sm">
          <PlusCircle className="w-4 h-4 mr-2" />
          Post New Job
        </Button>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Total"
          value={jobs.length}
          icon={Briefcase}
          iconColor="bg-slate-100 text-slate-600"
        />
        <StatCard
          label="Open"
          value={openJobs.length}
          icon={Briefcase}
          iconColor="bg-emerald-100 text-emerald-600"
        />
        <StatCard
          label="In Progress"
          value={inProgressJobs.length}
          icon={Briefcase}
          iconColor="bg-amber-100 text-amber-600"
        />
        <StatCard
          label="Total Proposals"
          value={totalProposals}
          icon={FileText}
          iconColor="bg-violet-100 text-violet-600"
        />
      </div>

      {/* Tabs */}
      <Tabs defaultValue="all">
        <TabsList className="bg-slate-100">
          <TabsTrigger value="all" className="text-sm">All ({jobs.length})</TabsTrigger>
          <TabsTrigger value="open" className="text-sm">Open ({openJobs.length})</TabsTrigger>
          <TabsTrigger value="in-progress" className="text-sm">In Progress ({inProgressJobs.length})</TabsTrigger>
          <TabsTrigger value="completed" className="text-sm">Completed ({completedJobs.length})</TabsTrigger>
          <TabsTrigger value="draft" className="text-sm">Draft ({draftJobs.length})</TabsTrigger>
        </TabsList>

        {['all', 'open', 'in-progress', 'completed', 'draft'].map(tab => (
          <TabsContent key={tab} value={tab}>
            {filteredJobs(tab).length === 0 ? (
              <Card className="border-slate-200">
                <CardContent className="p-12">
                  <EmptyState
                    icon={Briefcase}
                    title={tab === 'all' ? 'No jobs posted' : `No ${tab} jobs found`}
                    description={tab === 'all'
                      ? "You haven't posted any jobs yet. Post your first job to start receiving proposals."
                      : `You don't have any ${tab} jobs.`}
                    actionLabel="Post a Job"
                    onAction={() => navigate('jobs/post')}
                  />
                </CardContent>
              </Card>
            ) : (
              <div className="space-y-3">
                {filteredJobs(tab).map(job => (
                  <JobCard key={job.id} job={job} />
                ))}
              </div>
            )}
          </TabsContent>
        ))}
      </Tabs>
    </motion.div>
  )
}
