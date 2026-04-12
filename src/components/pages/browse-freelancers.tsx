'use client'

import { useEffect, useState, useCallback } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Input } from '@/components/ui/input'
import { Skeleton } from '@/components/ui/skeleton'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Switch } from '@/components/ui/switch'
import { authFetch } from '@/lib/api-fetch'
import {
  Search, Star, MapPin, DollarSign, ChevronLeft, ChevronRight,
  SlidersHorizontal, X, Grid3X3, List, Filter, Users,
  Clock, GraduationCap, MessageSquare, Eye, CircleCheck
} from 'lucide-react'

interface Freelancer {
  id: string
  displayName: string
  avatar: string | null
  title: string | null
  bio: string | null
  hourlyRate: number | null
  skills: string[]
  category: string
  location: string | null
  availabilityStatus: string
  rating: number
  totalReviews: number
  completedJobs: number
  experienceLevel: string
  englishLevel: string
  isStudent: boolean
  studentInstitution: string | null
  isFeatured: boolean
  createdAt: string
}

const CATEGORIES = [
  { value: '', label: 'All Categories' },
  { value: 'WEB_DEVELOPER', label: 'Web Development' },
  { value: 'DESIGNER', label: 'Graphic Design' },
  { value: 'VIDEO_EDITOR', label: 'Video Editing' },
  { value: 'CONTENT_CREATOR', label: 'Content Writing' },
  { value: 'OTHER', label: 'Other' },
]

const AVAILABILITY = [
  { value: '', label: 'Any Availability' },
  { value: 'AVAILABLE', label: 'Available' },
  { value: 'PARTIALLY_AVAILABLE', label: 'Partially Available' },
]

const RATING_FILTERS = [
  { value: '', label: 'Any Rating' },
  { value: '4', label: '4+ Stars' },
  { value: '3', label: '3+ Stars' },
  { value: '2', label: '2+ Stars' },
]

const SORT_OPTIONS = [
  { value: 'rating-desc', label: 'Highest Rated' },
  { value: 'createdAt-desc', label: 'Newest' },
  { value: 'hourlyRate-asc', label: 'Rate: Low to High' },
  { value: 'hourlyRate-desc', label: 'Rate: High to Low' },
  { value: 'completedJobs-desc', label: 'Most Jobs Completed' },
]

function formatExperience(level: string): string {
  if (!level) return ''
  return level.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatCategory(cat: string): string {
  if (!cat) return ''
  return cat.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatAvailability(status: string): string {
  if (!status) return ''
  return status.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

export default function BrowseFreelancers() {
  const { navigate, goBack } = useNavigationStore()
  const [freelancers, setFreelancers] = useState<Freelancer[]>([])
  const [loading, setLoading] = useState(true)
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('grid')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [total, setTotal] = useState(0)

  const [search, setSearch] = useState('')
  const [category, setCategory] = useState('')
  const [availability, setAvailability] = useState('')
  const [minRating, setMinRating] = useState('')
  const [sortBy, setSortBy] = useState('rating-desc')
  const [studentOnly, setStudentOnly] = useState(false)

  const fetchFreelancers = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', page.toString())
      params.set('limit', '12')
      params.set('sortBy', sortBy.split('-')[0])
      params.set('sortOrder', sortBy.split('-')[1] || 'desc')
      if (search) params.set('search', search)
      if (category) params.set('category', category)
      if (availability) params.set('availability', availability)
      if (minRating) params.set('minRating', minRating)
      if (studentOnly) params.set('isStudent', 'true')

      const res = await authFetch(`/api/freelancers?${params.toString()}`)
      const data = await res.json()
      if (data.freelancers) {
        setFreelancers(data.freelancers)
        setTotalPages(data.pagination?.totalPages || 1)
        setTotal(data.pagination?.total || 0)
      }
    } catch {
      setFreelancers([])
    } finally {
      setLoading(false)
    }
  }, [search, category, availability, minRating, sortBy, studentOnly, page])

  useEffect(() => {
    fetchFreelancers()
  }, [fetchFreelancers])

  const clearFilters = () => {
    setSearch('')
    setCategory('')
    setAvailability('')
    setMinRating('')
    setSortBy('rating-desc')
    setStudentOnly(false)
    setPage(1)
  }

  const hasActiveFilters = search || category || availability || minRating || studentOnly

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
                <Users className="w-5 h-5 text-emerald-600" />
                <h1 className="text-lg font-semibold">Browse Freelancers</h1>
              </div>
            </div>
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
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Search & Filters */}
        <Card className="mb-6 border-0 shadow-sm">
          <CardContent className="p-4 md:p-6 space-y-4">
            <div className="flex items-center gap-2 mb-2">
              <Filter className="w-4 h-4 text-muted-foreground" />
              <span className="text-sm font-medium">Filters</span>
            </div>

            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <Input
                placeholder="Search by name, title, skill, or bio..."
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

              <Select value={availability} onValueChange={(v) => { setAvailability(v === '__all__' ? '' : v); setPage(1) }}>
                <SelectTrigger>
                  <SelectValue placeholder="Any Availability" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="__all__">Any Availability</SelectItem>
                  {AVAILABILITY.filter(a => a.value).map(a => (
                    <SelectItem key={a.value} value={a.value}>{a.label}</SelectItem>
                  ))}
                </SelectContent>
              </Select>

              <Select value={minRating} onValueChange={(v) => { setMinRating(v === '__all__' ? '' : v); setPage(1) }}>
                <SelectTrigger>
                  <SelectValue placeholder="Any Rating" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="__all__">Any Rating</SelectItem>
                  {RATING_FILTERS.filter(r => r.value).map(r => (
                    <SelectItem key={r.value} value={r.value}>{r.label}</SelectItem>
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

            {/* Student toggle */}
            <div className="flex items-center justify-between p-3 rounded-xl bg-emerald-50 border border-emerald-100">
              <div className="flex items-center gap-2">
                <GraduationCap className="w-4 h-4 text-emerald-600" />
                <span className="text-sm font-medium text-emerald-900">Students Only</span>
                <span className="text-xs text-emerald-600">Show verified students</span>
              </div>
              <Switch checked={studentOnly} onCheckedChange={(v) => { setStudentOnly(v); setPage(1) }} />
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
                {availability && (
                  <Badge variant="secondary" className="gap-1">
                    {formatAvailability(availability)}
                    <X className="w-3 h-3 cursor-pointer" onClick={() => setAvailability('')} />
                  </Badge>
                )}
                {studentOnly && (
                  <Badge variant="secondary" className="gap-1">
                    Students
                    <X className="w-3 h-3 cursor-pointer" onClick={() => setStudentOnly(false)} />
                  </Badge>
                )}
                <Button variant="ghost" size="sm" onClick={clearFilters} className="text-xs h-6">
                  Clear all
                </Button>
              </div>
            )}
          </CardContent>
        </Card>

        {/* Results */}
        <div className="flex items-center justify-between mb-4">
          <p className="text-sm text-muted-foreground">
            {loading ? 'Searching...' : (
              <>
                <span className="font-semibold text-foreground">{total}</span> freelancers found
              </>
            )}
          </p>
        </div>

        {/* Freelancer Cards */}
        {loading ? (
          <div className={viewMode === 'grid' ? 'grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4' : 'space-y-4'}>
            {Array.from({ length: 6 }).map((_, i) => (
              <Card key={i} className="border-0 shadow-sm">
                <CardContent className="p-5">
                  <div className="flex items-center gap-4 mb-4">
                    <Skeleton className="w-14 h-14 rounded-full" />
                    <div className="space-y-2 flex-1">
                      <Skeleton className="h-4 w-3/4" />
                      <Skeleton className="h-3 w-1/2" />
                    </div>
                  </div>
                  <Skeleton className="h-3 w-full mb-2" />
                  <Skeleton className="h-3 w-2/3" />
                  <div className="flex gap-2 mt-4">
                    <Skeleton className="h-6 w-16 rounded-full" />
                    <Skeleton className="h-6 w-20 rounded-full" />
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
        ) : freelancers.length === 0 ? (
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="text-center py-20"
          >
            <div className="w-20 h-20 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-6">
              <Users className="w-10 h-10 text-gray-300" />
            </div>
            <h3 className="text-xl font-semibold text-gray-900 mb-2">No freelancers found</h3>
            <p className="text-muted-foreground mb-6 max-w-md mx-auto">
              Try adjusting your filters or search terms to find more freelancers.
            </p>
            <Button variant="outline" onClick={clearFilters}>
              Clear Filters
            </Button>
          </motion.div>
        ) : (
          <AnimatePresence mode="wait">
            <motion.div
              key={`fl-${viewMode}-${page}`}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
              className={viewMode === 'grid' ? 'grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4' : 'space-y-4'}
            >
              {freelancers.map((f) => (
                <Card
                  key={f.id}
                  className="group border-0 shadow-sm hover:shadow-md transition-all hover:-translate-y-0.5 h-full"
                >
                  <CardContent className="p-5 flex flex-col h-full">
                    {/* Avatar & Name */}
                    <div className="flex items-start gap-3 mb-3">
                      <div className="relative">
                        <Avatar className="w-14 h-14 border-2 border-gray-100">
                          <AvatarFallback className="bg-emerald-50 text-emerald-700 font-semibold text-sm">
                            {f.displayName.slice(0, 2).toUpperCase()}
                          </AvatarFallback>
                        </Avatar>
                        {/* Availability indicator */}
                        <div
                          className={`absolute -bottom-0.5 -right-0.5 w-4 h-4 rounded-full border-2 border-white ${
                            f.availabilityStatus === 'AVAILABLE'
                              ? 'bg-emerald-500'
                              : f.availabilityStatus === 'PARTIALLY_AVAILABLE'
                              ? 'bg-amber-500'
                              : 'bg-gray-400'
                          }`}
                        />
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <h3 className="font-semibold text-gray-900 truncate">{f.displayName}</h3>
                          {f.isFeatured && (
                            <Star className="w-3.5 h-3.5 fill-amber-400 text-amber-400 flex-shrink-0" />
                          )}
                        </div>
                        <p className="text-sm text-muted-foreground truncate">{f.title || 'Freelancer'}</p>
                      </div>
                    </div>

                    {/* Badges */}
                    <div className="flex flex-wrap gap-1.5 mb-3">
                      {f.category && (
                        <Badge variant="secondary" className="text-xs font-normal bg-emerald-50 text-emerald-700">
                          {formatCategory(f.category)}
                        </Badge>
                      )}
                      {f.isStudent && (
                        <Badge className="text-xs font-normal bg-violet-50 text-violet-700 gap-1">
                          <GraduationCap className="w-3 h-3" />
                          Student
                        </Badge>
                      )}
                      <Badge
                        variant="outline"
                        className={`text-xs font-normal ${
                          f.availabilityStatus === 'AVAILABLE'
                            ? 'text-emerald-600 border-emerald-200'
                            : 'text-amber-600 border-amber-200'
                        }`}
                      >
                        <CircleCheck className="w-3 h-3 mr-0.5" />
                        {f.availabilityStatus === 'AVAILABLE' ? 'Available' : f.availabilityStatus === 'PARTIALLY_AVAILABLE' ? 'Partially' : 'Unavailable'}
                      </Badge>
                    </div>

                    {/* Rating & Rate */}
                    <div className="flex items-center gap-4 mb-3 text-sm">
                      <div className="flex items-center gap-1">
                        <Star className="w-4 h-4 fill-amber-400 text-amber-400" />
                        <span className="font-semibold">{f.rating > 0 ? f.rating.toFixed(1) : 'New'}</span>
                        <span className="text-muted-foreground">({f.totalReviews})</span>
                      </div>
                      {f.hourlyRate && (
                        <div className="flex items-center gap-1 text-emerald-600 font-semibold">
                          <DollarSign className="w-3.5 h-3.5" />
                          {f.hourlyRate}/hr
                        </div>
                      )}
                    </div>

                    {/* Location */}
                    {f.location && (
                      <div className="flex items-center gap-1 text-xs text-muted-foreground mb-3">
                        <MapPin className="w-3 h-3" />
                        {f.location}
                      </div>
                    )}

                    {/* Skills */}
                    <div className="flex flex-wrap gap-1 mb-4 flex-1">
                      {(f.skills || []).slice(0, 4).map((skill: string) => (
                        <Badge key={skill} variant="outline" className="text-xs font-normal">
                          {skill}
                        </Badge>
                      ))}
                      {(f.skills || []).length > 4 && (
                        <Badge variant="outline" className="text-xs font-normal text-muted-foreground">
                          +{(f.skills || []).length - 4}
                        </Badge>
                      )}
                    </div>

                    {/* Bio preview */}
                    {f.bio && (
                      <p className="text-xs text-muted-foreground line-clamp-2 mb-4 leading-relaxed">{f.bio}</p>
                    )}

                    {/* Actions */}
                    <div className="flex gap-2 pt-3 border-t border-gray-100">
                      <Button
                        variant="outline"
                        className="flex-1 border-emerald-200 text-emerald-600 hover:bg-emerald-50 hover:text-emerald-700 text-sm"
                        onClick={() => navigate('freelancers/detail', { id: f.id })}
                      >
                        <Eye className="w-3.5 h-3.5 mr-1.5" />
                        View Profile
                      </Button>
                      <Button
                        variant="ghost"
                        size="sm"
                        className="text-muted-foreground hover:text-emerald-600"
                        onClick={() => navigate('messages')}
                      >
                        <MessageSquare className="w-4 h-4" />
                      </Button>
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
