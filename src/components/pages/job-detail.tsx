'use client'

import { useEffect, useState } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Textarea } from '@/components/ui/textarea'
import { Input } from '@/components/ui/input'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Skeleton } from '@/components/ui/skeleton'
import { Separator } from '@/components/ui/separator'
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'
import {
  ChevronLeft, Clock, DollarSign, MapPin, Star, Heart, Bookmark,
  Shield, AlertTriangle, FileText, Users, Send, CheckCircle,
  Briefcase, Calendar, Award, MessageSquare
} from 'lucide-react'

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
  attachments: any[]
  viewCount: number
  createdAt: string
  client: {
    id: string
    displayName: string
    avatar: string | null
    bio: string | null
    companyName: string | null
    location: string | null
    isVerified: boolean
    rating: number
    totalReviews: number
    jobsPosted: number
    memberSince: string | null
  }
}

interface Proposal {
  id: string
  coverLetter: string
  proposedAmount: number
  proposedDuration: string | null
  status: string
  createdAt: string
  freelancer: {
    id: string
    displayName: string
    avatar: string | null
    title: string | null
    rating: number
    completedJobs: number
  }
}

function formatDuration(duration: string | null): string {
  if (!duration) return 'Not specified'
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

function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric'
  })
}

export default function JobDetail() {
  const { pageParams, goBack, navigate } = useNavigationStore()
  const { toast } = useToast()
  const jobId = pageParams?.id

  const [job, setJob] = useState<Job | null>(null)
  const [proposals, setProposals] = useState<Proposal[]>([])
  const [loading, setLoading] = useState(true)
  const [proposalOpen, setProposalOpen] = useState(false)
  const [submitting, setSubmitting] = useState(false)
  const [saved, setSaved] = useState(false)
  const [saveLoading, setSaveLoading] = useState(false)
  const { user } = useAuthStore()

  const handleSaveJob = async () => {
    if (!user?.id) {
      toast({ title: 'Sign in required', description: 'Please sign in to save jobs.', variant: 'destructive' })
      return
    }
    setSaveLoading(true)
    try {
      const res = await authFetch('/api/saved-jobs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ jobId }),
      })
      if (res.ok) {
        setSaved(true)
        toast({ title: 'Job saved!', description: 'This job has been added to your saved jobs.' })
      } else {
        const data = await res.json()
        toast({ title: 'Error', description: data.error || 'Failed to save job', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to save job', variant: 'destructive' })
    }
    setSaveLoading(false)
  }

  // Proposal form
  const [coverLetter, setCoverLetter] = useState('')
  const [proposedAmount, setProposedAmount] = useState('')
  const [proposedDuration, setProposedDuration] = useState('')

  useEffect(() => {
    if (!jobId) return
    let cancelled = false
    const controller = new AbortController()

    authFetch(`/api/jobs/${jobId}`, { signal: controller.signal })
      .then(res => res.json())
      .then(data => {
        if (!cancelled && data.job) setJob(data.job)
      })
      .catch((err) => {
        if (err.name !== 'AbortError') console.error(err)
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })

    return () => {
      cancelled = true
      controller.abort()
    }
  }, [jobId])

  const fetchProposals = async () => {
    if (!jobId) return
    try {
      const res = await authFetch(`/api/jobs/${jobId}/proposals`)
      const data = await res.json()
      if (data.proposals) setProposals(data.proposals)
    } catch {
      // proposals only visible to job owner
    }
  }

  const submitProposal = async () => {
    if (!coverLetter.trim() || !proposedAmount) {
      toast({ title: 'Error', description: 'Cover letter and amount are required', variant: 'destructive' })
      return
    }
    setSubmitting(true)
    try {
      const res = await authFetch(`/api/jobs/${jobId}/proposals`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          coverLetter,
          proposedAmount: parseFloat(proposedAmount),
          proposedDuration: proposedDuration || undefined,
        }),
      })
      const data = await res.json()
      if (res.ok) {
        toast({ title: 'Success!', description: 'Proposal submitted successfully' })
        setProposalOpen(false)
        setCoverLetter('')
        setProposedAmount('')
        setProposedDuration('')
      } else {
        toast({ title: 'Error', description: data.error || 'Failed to submit proposal', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to submit proposal', variant: 'destructive' })
    } finally {
      setSubmitting(false)
    }
  }

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50/50 p-4 md:p-8">
        <div className="max-w-7xl mx-auto">
          <Skeleton className="h-8 w-40 mb-6" />
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-4">
              <Skeleton className="h-10 w-3/4" />
              <Skeleton className="h-20 w-full" />
              <Skeleton className="h-60 w-full" />
            </div>
            <div className="space-y-4">
              <Skeleton className="h-40 w-full" />
              <Skeleton className="h-60 w-full" />
            </div>
          </div>
        </div>
      </div>
    )
  }

  if (!job) {
    return (
      <div className="min-h-screen bg-gray-50/50 flex items-center justify-center">
        <div className="text-center">
          <h2 className="text-2xl font-bold text-gray-900 mb-2">Job not found</h2>
          <p className="text-muted-foreground mb-4">This job may have been removed or doesn&apos;t exist.</p>
          <Button onClick={() => goBack()}>Go Back</Button>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50/50">
      {/* Top Bar */}
      <div className="sticky top-0 z-40 bg-white border-b border-border">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-4 h-14">
            <button onClick={goBack} className="text-muted-foreground hover:text-foreground transition-colors">
              <ChevronLeft className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2">
              <Briefcase className="w-5 h-5 text-emerald-600" />
              <span className="text-sm text-muted-foreground">Job Details</span>
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Main Content */}
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className="lg:col-span-2 space-y-6"
          >
            {/* Title & Client */}
            <Card className="border-0 shadow-sm">
              <CardContent className="p-6">
                <div className="flex items-start gap-2 mb-4">
                  {job.isUrgent && (
                    <Badge variant="destructive" className="gap-1">
                      <AlertTriangle className="w-3 h-3" />
                      Urgent
                    </Badge>
                  )}
                  <Badge variant="secondary" className="bg-emerald-50 text-emerald-700">
                    {formatCategory(job.category)}
                  </Badge>
                </div>

                <h1 className="text-2xl md:text-3xl font-bold text-gray-900 mb-4">{job.title}</h1>

                {/* Client Info */}
                <div className="flex items-center gap-3 p-4 rounded-xl bg-gray-50">
                  <Avatar className="w-12 h-12">
                    <AvatarFallback className="bg-emerald-50 text-emerald-700 font-semibold text-sm">
                      {job.client.displayName.slice(0, 2).toUpperCase()}
                    </AvatarFallback>
                  </Avatar>
                  <div className="flex-1">
                    <div className="flex items-center gap-2">
                      <span className="font-semibold text-gray-900">{job.client.displayName}</span>
                      {job.client.isVerified && (
                        <span className="text-emerald-500 text-sm" title="Verified Client">✓ Verified</span>
                      )}
                    </div>
                    <div className="flex items-center gap-3 text-sm text-muted-foreground mt-0.5">
                      {job.client.location && (
                        <span className="flex items-center gap-1">
                          <MapPin className="w-3 h-3" />
                          {job.client.location}
                        </span>
                      )}
                      {job.client.memberSince && (
                        <span className="flex items-center gap-1">
                          <Calendar className="w-3 h-3" />
                          Member since {formatDate(job.client.memberSince)}
                        </span>
                      )}
                    </div>
                  </div>
                  {job.client.rating > 0 && (
                    <div className="flex items-center gap-1 px-3 py-1 rounded-lg bg-amber-50">
                      <Star className="w-4 h-4 fill-amber-400 text-amber-400" />
                      <span className="font-semibold text-sm">{job.client.rating.toFixed(1)}</span>
                      <span className="text-xs text-muted-foreground">({job.client.totalReviews})</span>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>

            {/* Job Details */}
            <Card className="border-0 shadow-sm">
              <CardContent className="p-6">
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                  <div className="p-3 rounded-xl bg-gray-50">
                    <div className="flex items-center gap-2 text-muted-foreground text-xs mb-1">
                      <Clock className="w-3.5 h-3.5" />
                      Posted
                    </div>
                    <div className="text-sm font-medium">{formatDate(job.createdAt)}</div>
                  </div>
                  <div className="p-3 rounded-xl bg-gray-50">
                    <div className="flex items-center gap-2 text-muted-foreground text-xs mb-1">
                      <DollarSign className="w-3.5 h-3.5" />
                      Budget
                    </div>
                    <div className="text-sm font-medium">
                      {job.budgetType === 'FIXED_PRICE'
                        ? `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''}`
                        : `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''}/hr`
                      }
                    </div>
                    <div className="text-xs text-muted-foreground">
                      {job.budgetType === 'FIXED_PRICE' ? 'Fixed Price' : 'Hourly'}
                    </div>
                  </div>
                  <div className="p-3 rounded-xl bg-gray-50">
                    <div className="flex items-center gap-2 text-muted-foreground text-xs mb-1">
                      <Calendar className="w-3.5 h-3.5" />
                      Duration
                    </div>
                    <div className="text-sm font-medium">{formatDuration(job.duration)}</div>
                  </div>
                  <div className="p-3 rounded-xl bg-gray-50">
                    <div className="flex items-center gap-2 text-muted-foreground text-xs mb-1">
                      <Award className="w-3.5 h-3.5" />
                      Experience
                    </div>
                    <div className="text-sm font-medium">{formatExperience(job.experienceLevel)}</div>
                  </div>
                </div>

                <Separator className="my-6" />

                <h2 className="text-lg font-semibold text-gray-900 mb-4">Job Description</h2>
                <div className="prose prose-sm max-w-none text-muted-foreground leading-relaxed whitespace-pre-wrap">
                  {job.description}
                </div>

                {/* Skills */}
                {job.skills && job.skills.length > 0 && (
                  <div className="mt-6">
                    <h3 className="text-sm font-semibold text-gray-900 mb-3">Skills Required</h3>
                    <div className="flex flex-wrap gap-2">
                      {job.skills.map((skill: string) => (
                        <Badge key={skill} variant="outline" className="font-normal">
                          {skill}
                        </Badge>
                      ))}
                    </div>
                  </div>
                )}

                {/* Attachments */}
                {job.attachments && job.attachments.length > 0 && (
                  <div className="mt-6">
                    <h3 className="text-sm font-semibold text-gray-900 mb-3">Attachments</h3>
                    <div className="space-y-2">
                      {job.attachments.map((att: any, i: number) => (
                        <div key={i} className="flex items-center gap-3 p-3 rounded-lg bg-gray-50">
                          <FileText className="w-4 h-4 text-muted-foreground" />
                          <span className="text-sm">{att.name || `File ${i + 1}`}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Proposals Section (visible to job owner) */}
            <Card className="border-0 shadow-sm">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <CardTitle className="text-lg flex items-center gap-2">
                    <Users className="w-5 h-5 text-emerald-600" />
                    Proposals ({job.proposalsCount})
                  </CardTitle>
                  <Button
                    variant="outline"
                    size="sm"
                    onClick={fetchProposals}
                  >
                    <MessageSquare className="w-4 h-4 mr-2" />
                    View Proposals
                  </Button>
                </div>
              </CardHeader>
              {proposals.length > 0 && (
                <CardContent className="pt-0 space-y-4">
                  {proposals.map((p) => (
                    <div key={p.id} className="p-4 rounded-xl border">
                      <div className="flex items-start gap-3 mb-3">
                        <Avatar className="w-10 h-10">
                          <AvatarFallback className="bg-emerald-50 text-emerald-700 font-semibold text-xs">
                            {p.freelancer.displayName.slice(0, 2).toUpperCase()}
                          </AvatarFallback>
                        </Avatar>
                        <div className="flex-1">
                          <div className="font-semibold text-sm">{p.freelancer.displayName}</div>
                          <div className="text-xs text-muted-foreground">{p.freelancer.title}</div>
                        </div>
                        <div className="text-right">
                          <div className="font-semibold text-emerald-600">${p.proposedAmount}</div>
                          <Badge variant="secondary" className="text-xs capitalize">
                            {p.status.toLowerCase()}
                          </Badge>
                        </div>
                      </div>
                      <p className="text-sm text-muted-foreground line-clamp-3">{p.coverLetter}</p>
                    </div>
                  ))}
                </CardContent>
              )}
            </Card>
          </motion.div>

          {/* Sidebar */}
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ delay: 0.1 }}
            className="space-y-4"
          >
            {/* CTA Card */}
            <Card className="border-0 shadow-sm sticky top-20">
              <CardContent className="p-6 space-y-4">
                <Dialog open={proposalOpen} onOpenChange={setProposalOpen}>
                  <DialogTrigger asChild>
                    <Button className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-6 text-base">
                      <Send className="w-4 h-4 mr-2" />
                      Submit Proposal
                    </Button>
                  </DialogTrigger>
                  <DialogContent className="sm:max-w-lg max-h-[90vh] overflow-y-auto">
                    <DialogHeader>
                      <DialogTitle>Submit Your Proposal</DialogTitle>
                    </DialogHeader>
                    <div className="space-y-4 pt-4">
                      <div>
                        <label className="text-sm font-medium mb-2 block">Cover Letter *</label>
                        <Textarea
                          placeholder="Tell the client why you're the best fit for this project. Share your relevant experience and approach..."
                          value={coverLetter}
                          onChange={(e) => setCoverLetter(e.target.value)}
                          rows={8}
                          maxLength={5000}
                          className="resize-none"
                        />
                        <div className="text-xs text-muted-foreground mt-1 text-right">
                          {coverLetter.length}/5000 characters
                        </div>
                      </div>

                      <div>
                        <label className="text-sm font-medium mb-2 block">Proposed Amount ($) *</label>
                        <Input
                          type="number"
                          placeholder="Enter your proposed amount"
                          value={proposedAmount}
                          onChange={(e) => setProposedAmount(e.target.value)}
                          min="1"
                        />
                      </div>

                      <div>
                        <label className="text-sm font-medium mb-2 block">Estimated Duration</label>
                        <Select value={proposedDuration} onValueChange={setProposedDuration}>
                          <SelectTrigger>
                            <SelectValue placeholder="Select duration" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="LESS_THAN_1_WEEK">Less than 1 week</SelectItem>
                            <SelectItem value="ONE_TO_4_WEEKS">1-4 weeks</SelectItem>
                            <SelectItem value="ONE_TO_3_MONTHS">1-3 months</SelectItem>
                            <SelectItem value="THREE_TO_6_MONTHS">3-6 months</SelectItem>
                            <SelectItem value="MORE_THAN_6_MONTHS">6+ months</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>

                      <div className="flex gap-3 pt-4">
                        <Button
                          variant="outline"
                          className="flex-1"
                          onClick={() => setProposalOpen(false)}
                        >
                          Cancel
                        </Button>
                        <Button
                          className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white"
                          onClick={submitProposal}
                          disabled={submitting}
                        >
                          {submitting ? 'Submitting...' : 'Submit Proposal'}
                        </Button>
                      </div>
                    </div>
                  </DialogContent>
                </Dialog>

                <Separator />

                {/* Job Summary */}
                <div className="space-y-3">
                  <h3 className="text-sm font-semibold text-gray-900">Job Details</h3>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Budget</span>
                    <span className="text-sm font-medium">
                      {job.budgetType === 'FIXED_PRICE'
                        ? `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''}`
                        : `$${job.budgetMin || 0}${job.budgetMax ? ` - $${job.budgetMax}` : ''}/hr`
                      }
                    </span>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Type</span>
                    <span className="text-sm font-medium">
                      {job.budgetType === 'FIXED_PRICE' ? 'Fixed Price' : 'Hourly'}
                    </span>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Duration</span>
                    <span className="text-sm font-medium">{formatDuration(job.duration)}</span>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Experience</span>
                    <span className="text-sm font-medium">{formatExperience(job.experienceLevel)}</span>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Proposals</span>
                    <span className="text-sm font-medium">{job.proposalsCount}</span>
                  </div>

                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Views</span>
                    <span className="text-sm font-medium">{job.viewCount}</span>
                  </div>
                </div>

                <Separator />

                {/* Skills */}
                {job.skills && job.skills.length > 0 && (
                  <div className="space-y-2">
                    <h3 className="text-sm font-semibold text-gray-900">Skills</h3>
                    <div className="flex flex-wrap gap-1.5">
                      {job.skills.map((skill: string) => (
                        <Badge key={skill} variant="outline" className="text-xs font-normal">
                          {skill}
                        </Badge>
                      ))}
                    </div>
                  </div>
                )}

                <Separator />

                {/* Actions */}
                <div className="space-y-2">
                  <Button
                    variant="outline"
                    className="w-full justify-start"
                    onClick={handleSaveJob}
                    disabled={saveLoading || saved}
                  >
                    <Heart className={`w-4 h-4 mr-2 ${saved ? 'fill-red-500 text-red-500' : ''}`} />
                    {saved ? 'Saved' : 'Save Job'}
                  </Button>
                  <Button
                    variant="ghost"
                    className="w-full justify-start text-muted-foreground hover:text-destructive"
                    size="sm"
                    onClick={() => navigate('reports', { jobId: job.id })}
                  >
                    <Shield className="w-4 h-4 mr-2" />
                    Report Job
                  </Button>
                </div>

                {/* About Client */}
                <Separator />
                <div className="space-y-3">
                  <h3 className="text-sm font-semibold text-gray-900">About the Client</h3>
                  <div className="flex items-center gap-2 text-sm text-muted-foreground">
                    <Briefcase className="w-4 h-4" />
                    <span>{job.client.jobsPosted || 0} jobs posted</span>
                  </div>
                  {job.client.companyName && (
                    <div className="text-sm text-muted-foreground">
                      Company: <span className="text-foreground">{job.client.companyName}</span>
                    </div>
                  )}
                  {job.client.bio && (
                    <p className="text-sm text-muted-foreground leading-relaxed">{job.client.bio}</p>
                  )}
                </div>
              </CardContent>
            </Card>
          </motion.div>
        </div>
      </div>
    </div>
  )
}
