'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency, timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import {
  Bookmark,
  Clock,
  Eye,
  Search,
  Trash2,
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'

interface SavedJob {
  id: string
  createdAt: string
  job: {
    id: string
    title: string
    status: string
    budgetType: string
    budgetMin: number | null
    budgetMax: number | null
    proposalsCount: number
    createdAt: string
    client: { id: string; displayName: string; avatar: string | null }
  }
}

export function SavedJobs() {
  const { navigate } = useNavigationStore()
  const { toast } = useToast()
  const [savedJobs, setSavedJobs] = useState<SavedJob[]>([])
  const [loading, setLoading] = useState(true)

  const fetchSavedJobs = useCallback(async () => {
    try {
      const res = await authFetch('/api/saved-jobs')
      const data = await res.json()
      if (data.savedJobs) setSavedJobs(data.savedJobs)
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => { fetchSavedJobs() }, [fetchSavedJobs])

  const handleRemove = async (jobId: string) => {
    try {
      const res = await authFetch(`/api/saved-jobs?jobId=${jobId}`, { method: 'DELETE' })
      if (res.ok) {
        toast({ title: 'Job removed', description: 'Job removed from saved list.' })
        setSavedJobs(prev => prev.filter(sj => sj.job.id !== jobId))
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to remove job', variant: 'destructive' })
    }
  }

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
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          {[1, 2, 3].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
        </div>
        {[1, 2, 3].map(i => <Skeleton key={i} className="h-24 w-full rounded-xl" />)}
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Saved Jobs</h1>
          <p className="text-slate-500 mt-1">
            {savedJobs.length} job{savedJobs.length !== 1 ? 's' : ''} bookmarked for later.
          </p>
        </div>
        <Button onClick={() => navigate('jobs')} variant="outline" className="border-slate-200">
          <Search className="w-4 h-4 mr-2" />
          Browse Jobs
        </Button>
      </div>

      {/* Summary Stats */}
      {savedJobs.length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <StatCard
            label="Saved Jobs"
            value={savedJobs.length}
            icon={Bookmark}
            iconColor="bg-emerald-100 text-emerald-600"
          />
          <StatCard
            label="Open Jobs"
            value={savedJobs.filter(sj => sj.job.status === 'open' || sj.job.status === 'OPEN').length}
            icon={Search}
            iconColor="bg-amber-100 text-amber-600"
          />
          <StatCard
            label="Recently Saved"
            value={savedJobs.filter(sj => {
              const diff = Date.now() - new Date(sj.createdAt).getTime()
              return diff < 7 * 24 * 60 * 60 * 1000
            }).length}
            icon={Clock}
            iconColor="bg-violet-100 text-violet-600"
            sub="Last 7 days"
          />
        </div>
      )}

      {savedJobs.length === 0 ? (
        <Card className="border-slate-200">
          <CardContent className="p-12">
            <EmptyState
              icon={Bookmark}
              title="No saved jobs"
              description="When you find jobs you're interested in, click the bookmark icon to save them here for easy access."
              actionLabel="Browse Jobs"
              onAction={() => navigate('jobs')}
            />
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-3">
          {savedJobs.map(savedJob => {
            const job = savedJob.job
            return (
              <Card key={savedJob.id} className="border-slate-200 hover:shadow-sm transition-shadow">
                <CardContent className="p-4">
                  <div className="flex items-start gap-4">
                    <div className="flex-1 min-w-0">
                      <div
                        className="text-sm font-semibold text-slate-900 hover:text-emerald-600 cursor-pointer truncate"
                        onClick={() => navigate('jobs/detail', { id: job.id })}
                      >
                        {job.title}
                      </div>
                      <div className="flex items-center gap-3 mt-1.5 flex-wrap">
                        <span className="text-sm font-medium text-emerald-600">
                          {job.budgetType === 'FIXED_PRICE'
                            ? `${formatCurrency(job.budgetMin || 0)} - ${formatCurrency(job.budgetMax || 0)}`
                            : 'Hourly'}
                        </span>
                        <span className="text-xs text-slate-400">•</span>
                        <span className="text-xs text-slate-500">{job.client?.displayName || 'Client'}</span>
                        <span className="text-xs text-slate-400">•</span>
                        <span className="text-xs text-slate-400 flex items-center gap-1">
                          <Clock className="w-3 h-3" />
                          {timeAgo(job.createdAt)}
                        </span>
                        <span className="text-xs text-slate-400">•</span>
                        <span className="text-xs text-slate-400">{job.proposalsCount} proposals</span>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => navigate('jobs/detail', { id: job.id })}
                        className="text-xs h-8 border-emerald-200 text-emerald-600 hover:bg-emerald-50"
                      >
                        <Eye className="w-3 h-3 mr-1" />
                        View
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => handleRemove(job.id)}
                        className="text-xs h-8 border-red-200 text-red-600 hover:bg-red-50"
                      >
                        <Trash2 className="w-3 h-3 mr-1" />
                        Remove
                      </Button>
                    </div>
                  </div>
                </CardContent>
              </Card>
            )
          })}
        </div>
      )}
    </motion.div>
  )
}
