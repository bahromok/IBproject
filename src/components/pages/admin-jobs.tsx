'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Card, CardContent } from '@/components/ui/card'
import { Input } from '@/components/ui/input'
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
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import {
  Search,
  Briefcase,
  Star,
  Eye,
  Trash2,
  MoreVertical,
  ChevronLeft,
  ChevronRight,
  User,
  Sparkles,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { EmptyState } from '@/components/shared/empty-state'
import { authFetch } from '@/lib/api-fetch'
import { toast } from 'sonner'

const STATUS_COLORS: Record<string, string> = {
  open: 'bg-emerald-100 text-emerald-700',
  OPEN: 'bg-emerald-100 text-emerald-700',
  DRAFT: 'bg-slate-100 text-slate-600',
  in_progress: 'bg-amber-100 text-amber-700',
  IN_PROGRESS: 'bg-amber-100 text-amber-700',
  completed: 'bg-emerald-100 text-emerald-600',
  COMPLETED: 'bg-emerald-100 text-emerald-600',
  cancelled: 'bg-red-100 text-red-700',
  CANCELLED: 'bg-red-100 text-red-700',
  deleted: 'bg-red-100 text-red-700',
  DELETED: 'bg-red-100 text-red-700',
}

const CATEGORY_LABELS: Record<string, string> = {
  DESIGNER: 'Designer',
  WEB_DEVELOPER: 'Web Developer',
  VIDEO_EDITOR: 'Video Editor',
  CONTENT_CREATOR: 'Content Creator',
  OTHER: 'Other',
}

export function AdminJobs() {
  const { navigate } = useNavigationStore()
  const [jobs, setJobs] = useState<any[]>([])
  const [featuredJobs, setFeaturedJobs] = useState<any[]>([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [statusFilter, setStatusFilter] = useState('all')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [deleteDialogOpen, setDeleteDialogOpen] = useState(false)
  const [deleteTarget, setDeleteTarget] = useState<string | null>(null)

  const fetchJobs = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', page.toString())
      params.set('limit', '20')
      if (search) params.set('search', search)
      if (categoryFilter !== 'all') params.set('category', categoryFilter)
      if (statusFilter !== 'all') params.set('status', statusFilter)

      const res = await authFetch(`/api/admin/jobs?${params.toString()}`)
      if (res.ok) {
        const data = await res.json()
        setJobs(data.jobs || [])
        setTotalPages(data.pagination?.totalPages || 1)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [page, search, categoryFilter, statusFilter])

  const fetchFeatured = useCallback(async () => {
    try {
      const res = await authFetch('/api/admin/jobs?isFeatured=true&limit=5')
      if (res.ok) {
        const data = await res.json()
        setFeaturedJobs(data.jobs || [])
      }
    } catch {
      // Silently handle
    }
  }, [])

  useEffect(() => { fetchJobs() }, [fetchJobs])
  useEffect(() => { fetchFeatured() }, [fetchFeatured])

  const handleAction = async (jobId: string, action: string) => {
    try {
      const res = await authFetch('/api/admin/jobs', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ jobId, action }),
      })
      if (res.ok) {
        toast.success(`Job ${action === 'feature' ? 'featured' : 'unfeatured'} successfully`)
        fetchJobs()
        fetchFeatured()
      } else {
        const data = await res.json()
        toast.error(data.error || 'Action failed')
      }
    } catch {
      toast.error('An error occurred')
    }
  }

  const handleDelete = async () => {
    if (!deleteTarget) return
    try {
      const res = await authFetch(`/api/admin/jobs?jobId=${deleteTarget}`, { method: 'DELETE' })
      if (res.ok) {
        toast.success('Job removed successfully')
        fetchJobs()
        fetchFeatured()
      } else {
        toast.error('Failed to remove job')
      }
    } catch {
      toast.error('An error occurred')
    } finally {
      setDeleteDialogOpen(false)
      setDeleteTarget(null)
    }
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-4"
    >
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Job Moderation</h1>
        <p className="text-sm text-slate-500 mt-0.5">Review and moderate job listings</p>
      </div>

      {/* Featured Jobs */}
      {featuredJobs.length > 0 && (
        <Card className="border-amber-200 bg-amber-50/30">
          <CardContent className="p-4">
            <div className="flex items-center gap-2 mb-3">
              <Sparkles className="w-4 h-4 text-amber-600" />
              <h2 className="text-sm font-semibold text-amber-900">Featured Jobs ({featuredJobs.length})</h2>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {featuredJobs.map((job: any) => (
                <div key={job.id} className="flex items-center gap-3 p-3 bg-white rounded-lg border border-amber-100">
                  <div className="w-8 h-8 rounded-lg bg-amber-100 flex items-center justify-center shrink-0">
                    <Briefcase className="w-4 h-4 text-amber-600" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-slate-900 truncate">{job.title}</p>
                    <p className="text-xs text-slate-500">{job.client?.clientProfile?.displayName || 'Client'}</p>
                  </div>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-7 text-xs text-amber-600"
                    onClick={() => handleAction(job.id, 'unfeature')}
                  >
                    Unfeature
                  </Button>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      {/* Filters */}
      <Card>
        <CardContent className="p-4">
          <div className="flex flex-col sm:flex-row gap-3">
            <div className="relative flex-1">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <Input
                placeholder="Search by title..."
                className="pl-9 h-9"
                value={search}
                onChange={e => { setSearch(e.target.value); setPage(1) }}
              />
            </div>
            <Select value={categoryFilter} onValueChange={v => { setCategoryFilter(v); setPage(1) }}>
              <SelectTrigger className="h-9 w-full sm:w-44">
                <SelectValue placeholder="Category" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Categories</SelectItem>
                {Object.entries(CATEGORY_LABELS).map(([k, v]) => (
                  <SelectItem key={k} value={k}>{v}</SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Select value={statusFilter} onValueChange={v => { setStatusFilter(v); setPage(1) }}>
              <SelectTrigger className="h-9 w-full sm:w-40">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Status</SelectItem>
                <SelectItem value="open">Open</SelectItem>
                <SelectItem value="DRAFT">Draft</SelectItem>
                <SelectItem value="in_progress">In Progress</SelectItem>
                <SelectItem value="completed">Completed</SelectItem>
                <SelectItem value="cancelled">Cancelled</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </CardContent>
      </Card>

      {/* Table */}
      <Card>
        <CardContent className="p-0">
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="pl-4">Title</TableHead>
                  <TableHead className="hidden md:table-cell">Client</TableHead>
                  <TableHead className="hidden lg:table-cell">Category</TableHead>
                  <TableHead className="hidden lg:table-cell">Budget</TableHead>
                  <TableHead>Status</TableHead>
                  <TableHead className="hidden sm:table-cell">Proposals</TableHead>
                  <TableHead className="text-right pr-4">Actions</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {loading ? (
                  Array.from({ length: 8 }).map((_, i) => (
                    <TableRow key={i}>
                      <TableCell className="pl-4"><Skeleton className="h-4 w-48" /></TableCell>
                      <TableCell className="hidden md:table-cell"><Skeleton className="h-4 w-28" /></TableCell>
                      <TableCell className="hidden lg:table-cell"><Skeleton className="h-5 w-20" /></TableCell>
                      <TableCell className="hidden lg:table-cell"><Skeleton className="h-4 w-20" /></TableCell>
                      <TableCell><Skeleton className="h-5 w-16" /></TableCell>
                      <TableCell className="hidden sm:table-cell"><Skeleton className="h-4 w-10" /></TableCell>
                      <TableCell className="text-right pr-4"><Skeleton className="h-8 w-8 ml-auto" /></TableCell>
                    </TableRow>
                  ))
                ) : jobs.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={7} className="py-12">
                      <EmptyState
                        icon={Briefcase}
                        title="No jobs found"
                        description={search || categoryFilter !== 'all' || statusFilter !== 'all'
                          ? 'Try adjusting your filters.'
                          : 'No jobs have been posted yet.'}
                      />
                    </TableCell>
                  </TableRow>
                ) : (
                  jobs.map((job: any) => (
                    <TableRow key={job.id} className="group">
                      <TableCell className="pl-4">
                        <div className="flex items-center gap-2">
                          {job.isFeatured && <Sparkles className="w-3.5 h-3.5 text-amber-500 shrink-0" />}
                          <span className="text-sm font-medium text-slate-900 truncate max-w-[200px]">
                            {job.title}
                          </span>
                        </div>
                      </TableCell>
                      <TableCell className="hidden md:table-cell text-sm text-slate-500">
                        {job.client?.clientProfile?.displayName || 'Unknown'}
                      </TableCell>
                      <TableCell className="hidden lg:table-cell">
                        <Badge variant="secondary" className="text-xs">
                          {CATEGORY_LABELS[job.category] || job.category}
                        </Badge>
                      </TableCell>
                      <TableCell className="hidden lg:table-cell text-sm text-slate-600">
                        {job.budgetMin && job.budgetMax
                          ? `$${job.budgetMin} - $${job.budgetMax}`
                          : job.budgetMin ? `$${job.budgetMin}+` : '-'}
                      </TableCell>
                      <TableCell>
                        <Badge variant="secondary" className={cn('text-xs', STATUS_COLORS[job.status] || 'bg-slate-100')}>
                          {job.status}
                        </Badge>
                      </TableCell>
                      <TableCell className="hidden sm:table-cell text-sm text-slate-500">
                        {job._count?.proposals || job.proposalCount || 0}
                      </TableCell>
                      <TableCell className="text-right pr-4">
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <Button variant="ghost" size="icon" className="h-8 w-8">
                              <MoreVertical className="w-4 h-4" />
                            </Button>
                          </DropdownMenuTrigger>
                          <DropdownMenuContent align="end">
                            <DropdownMenuItem onClick={() => navigate('jobs/detail', { id: job.id })}>
                              <Eye className="w-3.5 h-3.5 mr-2" /> View Job
                            </DropdownMenuItem>
                            <DropdownMenuItem onClick={() => navigate('freelancers/detail', { id: job.client?.id })}>
                              <User className="w-3.5 h-3.5 mr-2" /> View Client
                            </DropdownMenuItem>
                            <DropdownMenuSeparator />
                            {!job.isFeatured ? (
                              <DropdownMenuItem onClick={() => handleAction(job.id, 'feature')}>
                                <Star className="w-3.5 h-3.5 mr-2" /> Feature Job
                              </DropdownMenuItem>
                            ) : (
                              <DropdownMenuItem onClick={() => handleAction(job.id, 'unfeature')}>
                                <Star className="w-3.5 h-3.5 mr-2" /> Unfeature
                              </DropdownMenuItem>
                            )}
                            <DropdownMenuSeparator />
                            <DropdownMenuItem
                              onClick={() => {
                                setDeleteTarget(job.id)
                                setDeleteDialogOpen(true)
                              }}
                              className="text-red-600"
                            >
                              <Trash2 className="w-3.5 h-3.5 mr-2" /> Remove Job
                            </DropdownMenuItem>
                          </DropdownMenuContent>
                        </DropdownMenu>
                      </TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          </div>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100">
              <span className="text-sm text-slate-500">Page {page} of {totalPages}</span>
              <div className="flex items-center gap-1">
                <Button variant="outline" size="icon" className="h-8 w-8" onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page <= 1}>
                  <ChevronLeft className="w-4 h-4" />
                </Button>
                <Button variant="outline" size="icon" className="h-8 w-8" onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page >= totalPages}>
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Delete Dialog */}
      <AlertDialog open={deleteDialogOpen} onOpenChange={setDeleteDialogOpen}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Remove Job</AlertDialogTitle>
            <AlertDialogDescription>
              Are you sure you want to remove this job? This action will mark the job as deleted and cannot be undone.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction onClick={handleDelete} className="bg-red-600 hover:bg-red-700">
              Remove Job
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </motion.div>
  )
}
