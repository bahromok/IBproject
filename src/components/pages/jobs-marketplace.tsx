'use client'

import { useEffect, useState, useCallback } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Input } from '@/components/ui/input'
import { Skeleton } from '@/components/ui/skeleton'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { authFetch } from '@/lib/api-fetch'
import {
  Search, Heart, MapPin, Clock, DollarSign, Briefcase,
  SlidersHorizontal, X, Grid3X3, List, ChevronLeft, ChevronRight,
  Plus, ArrowUp, Flame, FileText
} from 'lucide-react'

const CATEGORIES = [
  { value: '', label: 'All Categories' },
  { value: 'WEB_DEVELOPER', label: 'Web Dev' },
  { value: 'MOBILE_DEVELOPER', label: 'Mobile App' },
  { value: 'DESIGNER', label: 'Design' },
  { value: 'VIDEO_EDITOR', label: 'Video Edit' },
  { value: 'CONTENT_CREATOR', label: 'Content' },
  { value: 'MARKETING', label: 'Marketing' },
  { value: 'WRITING', label: 'Writing' },
  { value: 'TRANSLATION', label: 'Translation' },
  { value: 'DATA_SCIENCE', label: 'Data Science' },
  { value: 'ADMIN_SUPPORT', label: 'Admin' },
  { value: 'OTHER', label: 'Other' },
]

const EXPERIENCE_LEVELS = [
  { value: '', label: 'Any Experience' },
  { value: 'BEGINNER', label: 'Beginner' },
  { value: 'INTERMEDIATE', label: 'Intermediate' },
  { value: 'ADVANCED', label: 'Advanced' },
  { value: 'EXPERT', label: 'Expert' },
]

const BUDGET_RANGES = [
  { value: '', label: 'Any Budget' },
  { value: '0-100', label: '$0 - $100' },
  { value: '100-500', label: '$100 - $500' },
  { value: '500-1000', label: '$500 - $1,000' },
  { value: '1000+', label: '$1,000+' },
]

const SORT_OPTIONS = [
  { value: 'createdAt-desc', label: 'Newest First' },
  { value: 'createdAt-asc', label: 'Oldest First' },
  { value: 'budget-desc', label: 'Budget: High to Low' },
  { value: 'budget-asc', label: 'Budget: Low to High' },
  { value: 'proposalsCount-desc', label: 'Most Proposals' },
]

interface Job {
  id: string
  title: string
  description: string
  category: string
  budgetType: string
  budgetMin: number | null
  budgetMax: number | null
  duration: string | null
  experienceLevel: string
  skills: string[]
  status: string
  proposalsCount: number
  isUrgent: boolean
  viewCount: number
  createdAt: string
  client: {
    id: string
    displayName: string
    avatar: string | null
    location: string | null
    isVerified: boolean
    rating: number
  }
}

function formatRelativeTime(dateStr: string): string {
  const now = new Date()
  const date = new Date(dateStr)
  const diffMs = now.getTime() - date.getTime()
  const diffMins = Math.floor(diffMs / 60000)
  const diffHours = Math.floor(diffMs / 3600000)
  const diffDays = Math.floor(diffMs / 86400000)
  const diffWeeks = Math.floor(diffDays / 7)

  if (diffMins < 1) return 'Just now'
  if (diffMins < 60) return `${diffMins}m ago`
  if (diffHours < 24) return `${diffHours}h ago`
  if (diffDays < 7) return `${diffDays}d ago`
  if (diffWeeks < 4) return `${diffWeeks}w ago`
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
}

function formatDuration(duration: string | null): string {
  if (!duration) return ''
  const map: Record<string, string> = {
    'LESS_THAN_1_WEEK': 'Less than 1 week',
    'ONE_TO_4_WEEKS': '1-4 weeks',
    'ONE_TO_3_MONTHS': '1-3 months',
    'THREE_TO_6_MONTHS': '3-6 months',
    'MORE_THAN_6_MONTHS': '6+ months',
  }
  return map[duration] || duration.replace(/_/g, ' ')
}

function formatExperience(level: string): string {
  return level.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatCategory(cat: string): string {
  return cat.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

export default function JobsMarketplace() {
  const { navigate, goBack } = useNavigationStore()
  const [jobs, setJobs] = useState<Job[]>([])
  const [loading, setLoading] = useState(true)
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('grid')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [total, setTotal] = useState(0)

  // Filters
  const [search, setSearch] = useState('')
  const [category, setCategory] = useState('')
  const [budgetRange, setBudgetRange] = useState('')
  const [experienceLevel, setExperienceLevel] = useState('')
  const [sortBy, setSortBy] = useState('createdAt-desc')

  const [savedJobs, setSavedJobs] = useState<Set<string>>(new Set())
  const [mobileFiltersOpen, setMobileFiltersOpen] = useState(false)

  const fetchJobs = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', page.toString())
      params.set('limit', '12')
      if (search) params.set('search', search)
      if (category) params.set('category', category)
      if (experienceLevel) params.set('experienceLevel', experienceLevel)

      if (budgetRange) {
        if (budgetRange === '0-100') {
          params.set('budgetMax', '100')
        } else if (budgetRange === '100-500') {
          params.set('budgetMin', '100')
          params.set('budgetMax', '500')
        } else if (budgetRange === '500-1000') {
          params.set('budgetMin', '500')
          params.set('budgetMax', '1000')
        } else if (budgetRange === '1000+') {
          params.set('budgetMin', '1000')
        }
      }

      const [sortField, sortOrder] = sortBy.split('-')
      params.set('sortBy', sortField)
      params.set('sortOrder', sortOrder)

      const res = await authFetch(`/api/jobs?${params.toString()}`)
      const data = await res.json()
      if (data.jobs) {
        setJobs(data.jobs)
        setTotalPages(data.pagination?.totalPages || 1)
        setTotal(data.pagination?.total || 0)
      }
    } catch {
      setJobs([])
    } finally {
      setLoading(false)
    }
  }, [search, category, budgetRange, experienceLevel, sortBy, page])

  useEffect(() => {
    fetchJobs()
  }, [fetchJobs])

  const clearFilters = () => {
    setSearch('')
    setCategory('')
    setBudgetRange('')
    setExperienceLevel('')
    setSortBy('createdAt-desc')
    setPage(1)
  }

  const toggleSaveJob = (jobId: string) => {
    setSavedJobs(prev => {
      const next = new Set(prev)
      if (next.has(jobId)) {
        next.delete(jobId)
      } else {
        next.add(jobId)
      }
      return next
    })
  }

  const hasActiveFilters = search || category || budgetRange || experienceLevel

  const filterBar = (
    <div className="space-y-4">
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
        <Input
          placeholder="Search by keyword, skill, or title..."
          value={search}
          onChange={(e) => { setSearch(e.target.value); setPage(1) }}
          className="pl-10"
        />
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <Select value={category} onValueChange={(v) => { setCategory(v === '__all__' ? '' : v); setPage(1) }}>
          <SelectTrigger>
            <SelectValue placeholder="All Categories" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="__all__">All Categories</SelectItem>
            {CATEGORIES.filter(c => c.value).map(c => (
              <SelectItem key={c.value} value={c.value}>{c.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>

        <Select value={budgetRange} onValueChange={(v) => { setBudgetRange(v === '__all__' ? '' : v); setPage(1) }}>
          <SelectTrigger>
            <SelectValue placeholder="Any Budget" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="__all__">Any Budget</SelectItem>
            {BUDGET_RANGES.filter(b => b.value).map(b => (
              <SelectItem key={b.value} value={b.value}>{b.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>

        <Select value={experienceLevel} onValueChange={(v) => { setExperienceLevel(v === '__all__' ? '' : v); setPage(1) }}>
          <SelectTrigger>
            <SelectValue placeholder="Any Experience" />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="__all__">Any Experience</SelectItem>
            {EXPERIENCE_LEVELS.filter(e => e.value).map(e => (
              <SelectItem key={e.value} value={e.value}>{e.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>

        <Select value={sortBy} onValueChange={setSortBy}>
          <SelectTrigger>
            <SelectValue placeholder="Sort by" />
          </SelectTrigger>
          <SelectContent>
            {SORT_OPTIONS.map(s => (
              <SelectItem key={s.value} value={s.value}>{s.label}</SelectItem>
            ))}
          </SelectContent>
        </Select>
      </div>

      {hasActiveFilters && (
        <div className="flex items-center gap-2">
          <span className="text-sm text-muted-foreground">Active filters:</span>
          {search && (
            <Badge variant="secondary" className="gap-1">
              &ldquo;{search}&rdquo;
              <X className="w-3 h-3 cursor-pointer" onClick={() => setSearch('')} />
            </Badge>
          )}
          {category && (
            <Badge variant="secondary" className="gap-1">
              {formatCategory(category)}
              <X className="w-3 h-3 cursor-pointer" onClick={() => setCategory('')} />
            </Badge>
          )}
          {experienceLevel && (
            <Badge variant="secondary" className="gap-1">
              {formatExperience(experienceLevel)}
              <X className="w-3 h-3 cursor-pointer" onClick={() => setExperienceLevel('')} />
            </Badge>
          )}
          <Button variant="ghost" size="sm" onClick={clearFilters} className="text-xs h-6">
            Clear all
          </Button>
        </div>
      )}
    </div>
  )

  return (
    <div className="min-h-screen bg-gray-50/50">
      {/* Top Nav Bar */}
      <div className="sticky top-0 z-40 bg-white border-b border-border">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-14">
            <div className="flex items-center gap-4">
              <button onClick={goBack} className="text-muted-foreground hover:text-foreground transition-colors">
                <ChevronLeft className="w-5 h-5" />
              </button>
              <div className="flex items-center gap-2">
                <Briefcase className="w-5 h-5 text-emerald-600" />
                <h1 className="text-lg font-semibold">Job Marketplace</h1>
              </div>
            </div>
            <div className="flex items-center gap-3">
              <div className="hidden sm:flex items-center gap-1 border rounded-lg p-1">
                <Button
                  variant={viewMode === 'grid' ? 'secondary' : 'ghost'}
                  size="sm"
                  className="h-8 w-8 p-0"
                  onClick={() => setViewMode('grid')}
                >
                  <Grid3X3 className="w-4 h-4" />
                </Button>
                <Button
                  variant={viewMode === 'list' ? 'secondary' : 'ghost'}
                  size="sm"
                  className="h-8 w-8 p-0"
                  onClick={() => setViewMode('list')}
                >
                  <List className="w-4 h-4" />
                </Button>
              </div>
              <Button
                className="bg-emerald-600 hover:bg-emerald-700 text-white"
                onClick={() => navigate('jobs/post')}
              >
                <Plus className="w-4 h-4 mr-2" />
                Post Job
              </Button>
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Search & Filters */}
        <Card className="mb-6 border-0 shadow-sm">
          <CardContent className="p-4 md:p-6">
            <div className="flex items-center gap-2 mb-4">
              <SlidersHorizontal className="w-4 h-4 text-muted-foreground" />
              <span className="text-sm font-medium">Filters</span>
            </div>
            {filterBar}
          </CardContent>
        </Card>

        {/* Results */}
        <div className="flex items-center justify-between mb-4">
          <p className="text-sm text-muted-foreground">
            {loading ? 'Searching...' : (
              <>
                <span className="font-semibold text-foreground">{total}</span> jobs found
              </>
            )}
          </p>
          {/* Mobile filter toggle */}
          <Button
            variant="outline"
            size="sm"
            className="md:hidden"
            onClick={() => setMobileFiltersOpen(!mobileFiltersOpen)}
          >
            <SlidersHorizontal className="w-4 h-4 mr-2" />
            Filters
          </Button>
        </div>

        {/* Job Cards */}
        {loading ? (
          <div className={viewMode === 'grid' ? 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4' : 'space-y-4'}>
            {Array.from({ length: 6 }).map((_, i) => (
              <Card key={i} className="border-0 shadow-sm">
                <CardContent className="p-5">
                  <Skeleton className="h-5 w-3/4 mb-3" />
                  <Skeleton className="h-4 w-1/2 mb-4" />
                  <div className="flex gap-2 mb-3">
                    <Skeleton className="h-6 w-20 rounded-full" />
                    <Skeleton className="h-6 w-24 rounded-full" />
                  </div>
                  <Skeleton className="h-4 w-full mb-2" />
                  <Skeleton className="h-4 w-2/3" />
                </CardContent>
              </Card>
            ))}
          </div>
        ) : jobs.length === 0 ? (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-center py-20"
          >
            <div className="w-20 h-20 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-6">
              <FileText className="w-10 h-10 text-gray-300" />
            </div>
            <h3 className="text-xl font-semibold text-gray-900 mb-2">No jobs found</h3>
            <p className="text-muted-foreground mb-6 max-w-md mx-auto">
              No jobs match your current criteria. Try adjusting your filters or search terms.
            </p>
            <Button variant="outline" onClick={clearFilters}>
              Clear Filters
            </Button>
          </motion.div>
        ) : (
          <AnimatePresence mode="wait">
            <motion.div
              key={`jobs-${viewMode}-${page}`}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
              className={viewMode === 'grid' ? 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4' : 'space-y-4'}
            >
              {jobs.map((job) => (
                <Card
                  key={job.id}
                  className="group border-0 shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 cursor-pointer h-full"
                >
                  <CardContent className="p-5 flex flex-col h-full">
                    <div className="flex items-start justify-between mb-3">
                      <div className="flex-1 min-w-0 mr-2">
                        <h3
                          className="font-semibold text-gray-900 group-hover:text-emerald-600 transition-colors line-clamp-2 leading-snug"
                          onClick={() => navigate('jobs/detail', { id: job.id })}
                        >
                          {job.title}
                        </h3>
                      </div>
                      <button
                        onClick={(e) => { e.stopPropagation(); toggleSaveJob(job.id) }}
                        className="flex-shrink-0 p-1 rounded-lg hover:bg-gray-100 transition-colors"
                      >
                        <Heart
                          className={`w-4 h-4 transition-colors ${savedJobs.has(job.id) ? 'fill-red-500 text-red-500' : 'text-gray-400'
                            }`}
                        />
                      </button>
                    </div>

                    <div className="flex items-center gap-2 text-xs text-muted-foreground mb-3">
                      <span className="font-medium text-foreground">{job.client.displayName}</span>
                      {job.client.isVerified && (
                        <span className="text-emerald-500" title="Verified">✓</span>
                      )}
                      <span>&middot;</span>
                      <Clock className="w-3 h-3" />
                      <span>{formatRelativeTime(job.createdAt)}</span>
                    </div>

                    <div className="flex flex-wrap gap-1.5 mb-3">
                      <Badge variant="secondary" className="text-xs font-normal bg-emerald-50 text-emerald-700">
                        {formatCategory(job.category)}
                      </Badge>
                      <Badge variant="secondary" className="text-xs font-normal">
                        {formatExperience(job.experienceLevel)}
                      </Badge>
                      {job.isUrgent && (
                        <Badge variant="destructive" className="text-xs gap-1">
                          <Flame className="w-3 h-3" />
                          Urgent
                        </Badge>
                      )}
                    </div>

                    <div className="flex items-center gap-2 mb-3">
                      <DollarSign className="w-4 h-4 text-emerald-600" />
                      <span className="text-sm font-semibold text-emerald-700">
                        {job.budgetType === 'FIXED_PRICE'
                          ? `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''} Fixed`
                          : `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''}/hr`
                        }
                      </span>
                    </div>

                    <div className="flex flex-wrap gap-1 mb-3 flex-1">
                      {(job.skills || []).slice(0, 4).map((skill: string) => (
                        <Badge key={skill} variant="outline" className="text-xs font-normal">
                          {skill}
                        </Badge>
                      ))}
                      {(job.skills || []).length > 4 && (
                        <Badge variant="outline" className="text-xs font-normal text-muted-foreground">
                          +{(job.skills || []).length - 4} more
                        </Badge>
                      )}
                    </div>

                    <div className="flex items-center justify-between text-xs text-muted-foreground pt-3 border-t border-gray-100">
                      <span>{job.proposalsCount} proposals</span>
                      <span className="flex items-center gap-1">
                        <ArrowUp className="w-3 h-3" />
                        {job.viewCount} views
                      </span>
                    </div>
                  </CardContent>
                </Card>
              ))}
            </motion.div>
          </AnimatePresence>
        )}

        {/* Pagination */}
        {totalPages > 1 && !loading && (
          <div className="flex items-center justify-center gap-2 mt-8">
            <Button
              variant="outline"
              size="sm"
              disabled={page <= 1}
              onClick={() => setPage(p => Math.max(1, p - 1))}
            >
              <ChevronLeft className="w-4 h-4" />
              Previous
            </Button>
            <div className="flex items-center gap-1">
              {Array.from({ length: Math.min(totalPages, 5) }, (_, i) => {
                const pageNum = page <= 3 ? i + 1 : page + i - 2
                if (pageNum > totalPages) return null
                return (
                  <Button
                    key={pageNum}
                    variant={pageNum === page ? 'default' : 'outline'}
                    size="sm"
                    className={pageNum === page ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'w-9'}
                    onClick={() => setPage(pageNum)}
                  >
                    {pageNum}
                  </Button>
                )
              })}
            </div>
            <Button
              variant="outline"
              size="sm"
              disabled={page >= totalPages}
              onClick={() => setPage(p => Math.min(totalPages, p + 1))}
            >
              Next
              <ChevronRight className="w-4 h-4" />
            </Button>
          </div>
        )}
      </div>
    </div>
  )
}
