'use client'

import { useEffect, useState } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Skeleton } from '@/components/ui/skeleton'
import { Separator } from '@/components/ui/separator'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import {
  ChevronLeft, Star, MapPin, DollarSign, Globe, Clock,
  Briefcase, GraduationCap, Award, Shield, MessageSquare,
  ExternalLink, CheckCircle, BarChart3, Users, Calendar,
  CircleCheck, Eye
} from 'lucide-react'

interface FreelancerProfile {
  id: string
  displayName: string
  avatar: string | null
  title: string | null
  bio: string | null
  hourlyRate: number | null
  skills: string[]
  portfolio: any[]
  category: string
  location: string | null
  timezone: string | null
  availabilityStatus: string
  rating: number
  totalReviews: number
  completedJobs: number
  experienceLevel: string
  englishLevel: string
  isStudent: boolean
  studentInstitution: string | null
  studentMajor: string | null
  expectedGraduation: string | null
  isVerified: boolean
  isFeatured: boolean
  stats: {
    activeContracts: number
    completedContracts: number
  }
  reviews: {
    id: string
    rating: number
    quality: number
    communication: number
    professionalism: number
    comment: string | null
    createdAt: string
    reviewer: {
      displayName: string
      avatar: string | null
    }
    contract: {
      id: string
      title: string
    }
  }[]
}

function formatExperience(level: string): string {
  if (!level) return 'Not specified'
  return level.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatCategory(cat: string): string {
  if (!cat) return ''
  return cat.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatEnglish(level: string): string {
  if (!level) return ''
  return level.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())
}

function formatDate(dateStr: string): string {
  return new Date(dateStr).toLocaleDateString('en-US', {
    year: 'numeric', month: 'long', day: 'numeric'
  })
}

function StarRating({ rating, size = 'sm' }: { rating: number; size?: 'sm' | 'md' }) {
  const sz = size === 'md' ? 'w-5 h-5' : 'w-4 h-4'
  return (
    <div className="flex items-center gap-0.5">
      {[1, 2, 3, 4, 5].map(star => (
        <Star
          key={star}
          className={`${sz} ${star <= Math.round(rating)
              ? 'fill-amber-400 text-amber-400'
              : 'fill-gray-200 text-gray-200'
            }`}
        />
      ))}
    </div>
  )
}

export default function FreelancerProfile() {
  const { pageParams, goBack, navigate } = useNavigationStore()
  const profileId = pageParams?.id

  const [profile, setProfile] = useState<FreelancerProfile | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!profileId) return
    let cancelled = false
    const controller = new AbortController()

    fetch(`/api/freelancers/${profileId}`, { signal: controller.signal })
      .then(res => res.json())
      .then(data => {
        if (!cancelled && data.freelancer) setProfile(data.freelancer)
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
  }, [profileId])

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-50/50 p-4 md:p-8">
        <div className="max-w-6xl mx-auto">
          <Skeleton className="h-8 w-40 mb-6" />
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-2 space-y-4">
              <Skeleton className="h-48 w-full rounded-xl" />
              <Skeleton className="h-60 w-full" />
            </div>
            <div className="space-y-4">
              <Skeleton className="h-40 w-full rounded-xl" />
              <Skeleton className="h-60 w-full rounded-xl" />
            </div>
          </div>
        </div>
      </div>
    )
  }

  if (!profile) {
    return (
      <div className="min-h-screen bg-gray-50/50 flex items-center justify-center">
        <div className="text-center">
          <h2 className="text-2xl font-bold text-gray-900 mb-2">Freelancer not found</h2>
          <p className="text-muted-foreground mb-4">This profile may have been removed or doesn&apos;t exist.</p>
          <Button onClick={goBack}>Go Back</Button>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50/50">
      {/* Top Bar */}
      <div className="sticky top-0 z-40 bg-white border-b border-border">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center gap-4 h-14">
            <button onClick={goBack} className="text-muted-foreground hover:text-foreground transition-colors">
              <ChevronLeft className="w-5 h-5" />
            </button>
            <span className="text-sm text-muted-foreground">Freelancer Profile</span>
          </div>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Main Content */}
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className="lg:col-span-2 space-y-6"
          >
            {/* Profile Header */}
            <Card className="border-0 shadow-sm overflow-hidden">
              <div className="bg-gradient-to-r from-emerald-500 to-teal-600 h-24 md:h-32" />
              <CardContent className="p-6 relative">
                <div className="flex flex-col sm:flex-row gap-4 -mt-12 sm:-mt-10">
                  <Avatar className="w-24 h-24 border-4 border-white shadow-lg">
                    <AvatarFallback className="bg-emerald-100 text-emerald-700 font-bold text-xl">
                      {profile.displayName.slice(0, 2).toUpperCase()}
                    </AvatarFallback>
                  </Avatar>
                  <div className="flex-1 pt-2 sm:pt-10">
                    <div className="flex flex-wrap items-center gap-2">
                      <h1 className="text-2xl font-bold text-gray-900">{profile.displayName}</h1>
                      {profile.isVerified && (
                        <CheckCircle className="w-5 h-5 text-emerald-500" />
                      )}
                      {profile.isFeatured && (
                        <Star className="w-5 h-5 fill-amber-400 text-amber-400" />
                      )}
                    </div>
                    <p className="text-muted-foreground">{profile.title || 'Freelancer'}</p>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-4 mt-4 text-sm text-muted-foreground">
                  {profile.location && (
                    <span className="flex items-center gap-1">
                      <MapPin className="w-4 h-4" />
                      {profile.location}
                    </span>
                  )}
                  {profile.timezone && (
                    <span className="flex items-center gap-1">
                      <Globe className="w-4 h-4" />
                      {profile.timezone}
                    </span>
                  )}
                  <div
                    className={`flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium ${profile.availabilityStatus === 'AVAILABLE'
                        ? 'bg-emerald-50 text-emerald-700'
                        : profile.availabilityStatus === 'PARTIALLY_AVAILABLE'
                          ? 'bg-amber-50 text-amber-700'
                          : 'bg-gray-100 text-gray-600'
                      }`}
                  >
                    <CircleCheck className="w-3.5 h-3.5" />
                    {profile.availabilityStatus === 'AVAILABLE'
                      ? 'Available now'
                      : profile.availabilityStatus === 'PARTIALLY_AVAILABLE'
                        ? 'Partially available'
                        : 'Not available'}
                  </div>
                </div>

                {/* Action buttons */}
                <div className="flex flex-wrap gap-3 mt-6">
                  <Button className="bg-emerald-600 hover:bg-emerald-700 text-white" onClick={() => navigate('jobs/post')}>
                    <Briefcase className="w-4 h-4 mr-2" />
                    Hire Me
                  </Button>
                  <Button variant="outline" onClick={() => navigate('messages')}>
                    <MessageSquare className="w-4 h-4 mr-2" />
                    Message
                  </Button>
                  <Button variant="ghost" size="sm" className="text-muted-foreground hover:text-destructive ml-auto" onClick={() => navigate('reports', { userId: profile.id })}>
                    <Shield className="w-4 h-4 mr-1.5" />
                    Report
                  </Button>
                </div>
              </CardContent>
            </Card>

            {/* Tabs: About, Portfolio, Reviews */}
            <Tabs defaultValue="about" className="w-full">
              <TabsList className="bg-white shadow-sm border rounded-lg w-full justify-start p-0 h-auto">
                <TabsTrigger value="about" className="rounded-lg px-6 py-3 data-[state=active]:bg-emerald-50 data-[state=active]:text-emerald-700">About</TabsTrigger>
                <TabsTrigger value="portfolio" className="rounded-lg px-6 py-3 data-[state=active]:bg-emerald-50 data-[state=active]:text-emerald-700">Portfolio</TabsTrigger>
                <TabsTrigger value="reviews" className="rounded-lg px-6 py-3 data-[state=active]:bg-emerald-50 data-[state=active]:text-emerald-700">Reviews ({profile.totalReviews})</TabsTrigger>
              </TabsList>

              <TabsContent value="about" className="mt-6 space-y-6">
                {/* Bio */}
                {profile.bio && (
                  <Card className="border-0 shadow-sm">
                    <CardContent className="p-6">
                      <h3 className="font-semibold text-gray-900 mb-3">About Me</h3>
                      <div className="text-sm text-muted-foreground leading-relaxed whitespace-pre-wrap">
                        {profile.bio}
                      </div>
                    </CardContent>
                  </Card>
                )}

                {/* Student Info */}
                {profile.isStudent && (
                  <Card className="border-0 shadow-sm">
                    <CardContent className="p-6">
                      <div className="flex items-center gap-2 mb-4">
                        <GraduationCap className="w-5 h-5 text-violet-600" />
                        <h3 className="font-semibold text-gray-900">Student Information</h3>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {profile.studentInstitution && (
                          <div>
                            <div className="text-xs text-muted-foreground">Institution</div>
                            <div className="text-sm font-medium">{profile.studentInstitution}</div>
                          </div>
                        )}
                        {profile.studentMajor && (
                          <div>
                            <div className="text-xs text-muted-foreground">Major</div>
                            <div className="text-sm font-medium">{profile.studentMajor}</div>
                          </div>
                        )}
                        {profile.expectedGraduation && (
                          <div>
                            <div className="text-xs text-muted-foreground">Expected Graduation</div>
                            <div className="text-sm font-medium">{profile.expectedGraduation}</div>
                          </div>
                        )}
                      </div>
                    </CardContent>
                  </Card>
                )}

                {/* Skills */}
                {profile.skills && profile.skills.length > 0 && (
                  <Card className="border-0 shadow-sm">
                    <CardContent className="p-6">
                      <h3 className="font-semibold text-gray-900 mb-3">Skills</h3>
                      <div className="flex flex-wrap gap-2">
                        {profile.skills.map((skill: string) => (
                          <Badge key={skill} variant="secondary" className="font-normal">
                            {skill}
                          </Badge>
                        ))}
                      </div>
                    </CardContent>
                  </Card>
                )}

                {/* Details */}
                <Card className="border-0 shadow-sm">
                  <CardContent className="p-6">
                    <h3 className="font-semibold text-gray-900 mb-4">Details</h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-lg bg-gray-50 flex items-center justify-center">
                          <Award className="w-5 h-5 text-muted-foreground" />
                        </div>
                        <div>
                          <div className="text-xs text-muted-foreground">Experience Level</div>
                          <div className="text-sm font-medium">{formatExperience(profile.experienceLevel)}</div>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-lg bg-gray-50 flex items-center justify-center">
                          <Globe className="w-5 h-5 text-muted-foreground" />
                        </div>
                        <div>
                          <div className="text-xs text-muted-foreground">English Level</div>
                          <div className="text-sm font-medium">{formatEnglish(profile.englishLevel)}</div>
                        </div>
                      </div>
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-lg bg-gray-50 flex items-center justify-center">
                          <DollarSign className="w-5 h-5 text-muted-foreground" />
                        </div>
                        <div>
                          <div className="text-xs text-muted-foreground">Hourly Rate</div>
                          <div className="text-sm font-medium">{profile.hourlyRate ? `$${profile.hourlyRate}/hr` : 'Not set'}</div>
                        </div>
                      </div>
                      {profile.category && (
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-lg bg-gray-50 flex items-center justify-center">
                            <Briefcase className="w-5 h-5 text-muted-foreground" />
                          </div>
                          <div>
                            <div className="text-xs text-muted-foreground">Category</div>
                            <div className="text-sm font-medium">{formatCategory(profile.category)}</div>
                          </div>
                        </div>
                      )}
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="portfolio" className="mt-6">
                <Card className="border-0 shadow-sm">
                  <CardContent className="p-6">
                    {profile.portfolio && profile.portfolio.length > 0 ? (
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        {profile.portfolio.map((item: any, i: number) => (
                          <div key={i} className="group rounded-xl border overflow-hidden hover:shadow-md transition-shadow">
                            <div className="aspect-video bg-gray-100 flex items-center justify-center">
                              <Eye className="w-8 h-8 text-gray-300" />
                            </div>
                            <div className="p-4">
                              <h4 className="font-medium text-sm">{item.title || `Project ${i + 1}`}</h4>
                              {item.description && (
                                <p className="text-xs text-muted-foreground mt-1 line-clamp-2">{item.description}</p>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="text-center py-12">
                        <div className="w-16 h-16 rounded-full bg-gray-100 flex items-center justify-center mx-auto mb-4">
                          <ExternalLink className="w-8 h-8 text-gray-300" />
                        </div>
                        <h3 className="font-medium text-gray-900 mb-1">No portfolio items yet</h3>
                        <p className="text-sm text-muted-foreground">
                          This freelancer hasn&apos;t added any portfolio items.
                        </p>
                      </div>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="reviews" className="mt-6 space-y-4">
                {profile.reviews && profile.reviews.length > 0 ? (
                  profile.reviews.map((review) => (
                    <Card key={review.id} className="border-0 shadow-sm">
                      <CardContent className="p-6">
                        <div className="flex items-start justify-between mb-3">
                          <div className="flex items-center gap-3">
                            <Avatar className="w-10 h-10">
                              <AvatarFallback className="bg-gray-100 text-gray-600 font-semibold text-xs">
                                {review.reviewer.displayName.slice(0, 2).toUpperCase()}
                              </AvatarFallback>
                            </Avatar>
                            <div>
                              <div className="font-medium text-sm">{review.reviewer.displayName}</div>
                              <div className="text-xs text-muted-foreground">
                                {formatDate(review.createdAt)}
                                {review.contract?.title && (
                                  <span className="ml-2">&middot; {review.contract.title}</span>
                                )}
                              </div>
                            </div>
                          </div>
                          <div className="flex items-center gap-1">
                            <Star className="w-4 h-4 fill-amber-400 text-amber-400" />
                            <span className="font-semibold text-sm">{review.rating.toFixed(1)}</span>
                          </div>
                        </div>

                        <div className="grid grid-cols-3 gap-3 mb-3">
                          <div className="text-center p-2 rounded-lg bg-gray-50">
                            <div className="text-xs text-muted-foreground">Quality</div>
                            <div className="text-sm font-semibold">{review.quality?.toFixed(1) || '-'}</div>
                          </div>
                          <div className="text-center p-2 rounded-lg bg-gray-50">
                            <div className="text-xs text-muted-foreground">Communication</div>
                            <div className="text-sm font-semibold">{review.communication?.toFixed(1) || '-'}</div>
                          </div>
                          <div className="text-center p-2 rounded-lg bg-gray-50">
                            <div className="text-xs text-muted-foreground">Professionalism</div>
                            <div className="text-sm font-semibold">{review.professionalism?.toFixed(1) || '-'}</div>
                          </div>
                        </div>

                        {review.comment && (
                          <p className="text-sm text-muted-foreground leading-relaxed italic">
                            &ldquo;{review.comment}&rdquo;
                          </p>
                        )}
                      </CardContent>
                    </Card>
                  ))
                ) : (
                  <Card className="border-0 shadow-sm">
                    <CardContent className="p-6 text-center py-12">
                      <Star className="w-10 h-10 text-gray-300 mx-auto mb-4" />
                      <h3 className="font-medium text-gray-900 mb-1">No reviews yet</h3>
                      <p className="text-sm text-muted-foreground">
                        This freelancer hasn&apos;t received any reviews yet.
                      </p>
                    </CardContent>
                  </Card>
                )}
              </TabsContent>
            </Tabs>
          </motion.div>

          {/* Sidebar */}
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ delay: 0.1 }}
            className="space-y-4"
          >
            {/* Stats Card */}
            <Card className="border-0 shadow-sm">
              <CardContent className="p-6">
                <h3 className="font-semibold text-gray-900 mb-4 flex items-center gap-2">
                  <BarChart3 className="w-5 h-5 text-emerald-600" />
                  Stats
                </h3>
                <div className="space-y-4">
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground">Rating</span>
                    <div className="flex items-center gap-1.5">
                      <StarRating rating={profile.rating} />
                      <span className="text-sm font-semibold">{profile.rating > 0 ? profile.rating.toFixed(1) : '-'}</span>
                    </div>
                  </div>
                  <Separator />
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground flex items-center gap-1.5">
                      <Users className="w-3.5 h-3.5" />
                      Reviews
                    </span>
                    <span className="text-sm font-semibold">{profile.totalReviews}</span>
                  </div>
                  <Separator />
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground flex items-center gap-1.5">
                      <CheckCircle className="w-3.5 h-3.5" />
                      Completed Jobs
                    </span>
                    <span className="text-sm font-semibold">{profile.completedJobs}</span>
                  </div>
                  <Separator />
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground flex items-center gap-1.5">
                      <Briefcase className="w-3.5 h-3.5" />
                      Active Contracts
                    </span>
                    <span className="text-sm font-semibold">{profile.stats?.activeContracts || 0}</span>
                  </div>
                  <Separator />
                  <div className="flex items-center justify-between">
                    <span className="text-sm text-muted-foreground flex items-center gap-1.5">
                      <DollarSign className="w-3.5 h-3.5" />
                      Hourly Rate
                    </span>
                    <span className="text-sm font-semibold text-emerald-600">
                      {profile.hourlyRate ? `$${profile.hourlyRate}/hr` : '-'}
                    </span>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Quick Info */}
            <Card className="border-0 shadow-sm">
              <CardContent className="p-6">
                <h3 className="font-semibold text-gray-900 mb-4">Quick Info</h3>
                <div className="space-y-3">
                  {profile.location && (
                    <div className="flex items-center gap-3 text-sm">
                      <MapPin className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                      <span>{profile.location}</span>
                    </div>
                  )}
                  {profile.timezone && (
                    <div className="flex items-center gap-3 text-sm">
                      <Clock className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                      <span>{profile.timezone}</span>
                    </div>
                  )}
                  {profile.category && (
                    <div className="flex items-center gap-3 text-sm">
                      <Briefcase className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                      <span>{formatCategory(profile.category)}</span>
                    </div>
                  )}
                  <div className="flex items-center gap-3 text-sm">
                    <Award className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                    <span>{formatExperience(profile.experienceLevel)}</span>
                  </div>
                  {profile.englishLevel && (
                    <div className="flex items-center gap-3 text-sm">
                      <Globe className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                      <span>{formatEnglish(profile.englishLevel)}</span>
                    </div>
                  )}
                </div>
              </CardContent>
            </Card>

            {/* Skills Summary */}
            {profile.skills && profile.skills.length > 0 && (
              <Card className="border-0 shadow-sm">
                <CardContent className="p-6">
                  <h3 className="font-semibold text-gray-900 mb-3">Skills</h3>
                  <div className="flex flex-wrap gap-1.5">
                    {profile.skills.map((skill: string) => (
                      <Badge key={skill} variant="outline" className="font-normal text-xs">
                        {skill}
                      </Badge>
                    ))}
                  </div>
                </CardContent>
              </Card>
            )}

            {/* CTA */}
            <Card className="border-2 border-emerald-200 bg-emerald-50/50">
              <CardContent className="p-6 text-center space-y-4">
                <h3 className="font-semibold text-gray-900">Interested in working together?</h3>
                <Button className="w-full bg-emerald-600 hover:bg-emerald-700 text-white" onClick={() => navigate('jobs/post')}>
                  <Briefcase className="w-4 h-4 mr-2" />
                  Invite to Job
                </Button>
                <Button variant="outline" className="w-full" onClick={() => navigate('messages')}>
                  <MessageSquare className="w-4 h-4 mr-2" />
                  Send Message
                </Button>
              </CardContent>
            </Card>
          </motion.div>
        </div>
      </div>
    </div>
  )
}
