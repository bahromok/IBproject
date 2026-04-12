'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { Card, CardContent } from '@/components/ui/card'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Skeleton } from '@/components/ui/skeleton'
import { Progress } from '@/components/ui/progress'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import { Star, MessageSquare, User } from 'lucide-react'

interface Review {
  id: string
  rating: number
  quality: number
  communication: number
  professionalism: number
  comment: string | null
  createdAt: string
  reviewer: { id: string; displayName: string; avatar: string | null }
  contract: { id: string; title: string }
}

interface ReviewStats {
  total: number
  averageRating: number
}

function StarRating({ rating }: { rating: number }) {
  return (
    <div className="flex items-center gap-0.5">
      {[1, 2, 3, 4, 5].map(star => (
        <Star
          key={star}
          className={`w-4 h-4 ${star <= rating ? 'text-amber-400 fill-amber-400' : 'text-slate-200'}`}
        />
      ))}
    </div>
  )
}

export function ReviewsPage() {
  const { user } = useAuthStore()
  const [reviews, setReviews] = useState<Review[]>([])
  const [stats, setStats] = useState<ReviewStats>({ total: 0, averageRating: 0 })
  const [loading, setLoading] = useState(true)

  const fetchReviews = useCallback(async () => {
    if (!user?.id) return
    try {
      const res = await authFetch(`/api/reviews?userId=${user.id}&limit=50`)
      const data = await res.json()
      if (data.reviews) setReviews(data.reviews)

      // Calculate stats from real data if API doesn't return stats
      const fetchedReviews = data.reviews || []
      if (fetchedReviews.length > 0) {
        const avgRating = fetchedReviews.reduce((sum: number, r: Review) => sum + r.rating, 0) / fetchedReviews.length
        setStats({ total: fetchedReviews.length, averageRating: Math.round(avgRating * 10) / 10 })
      } else if (data.stats) {
        setStats(data.stats)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [user?.id])

  useEffect(() => { fetchReviews() }, [fetchReviews])

  const distribution = [5, 4, 3, 2, 1].map(star => ({
    star,
    count: reviews.filter(r => Math.floor(r.rating) === star).length,
    percentage: reviews.length > 0 ? (reviews.filter(r => Math.floor(r.rating) === star).length / reviews.length) * 100 : 0,
  }))

  const ReviewCard = ({ review }: { review: Review }) => (
    <Card className="border-slate-200">
      <CardContent className="p-4">
        <div className="flex items-start gap-3">
          <Avatar className="h-10 w-10 border border-slate-200 shrink-0">
            <AvatarFallback className="bg-emerald-100 text-emerald-700 text-xs font-semibold">
              {review.reviewer?.displayName?.[0] || 'U'}
            </AvatarFallback>
          </Avatar>
          <div className="flex-1 min-w-0">
            <div className="flex items-center justify-between">
              <div>
                <span className="text-sm font-medium text-slate-900">{review.reviewer?.displayName || 'Anonymous'}</span>
                <div className="text-xs text-slate-400 mt-0.5">
                  {timeAgo(review.createdAt)}
                </div>
              </div>
              <StarRating rating={review.rating} />
            </div>
            {review.contract && (
              <div className="text-xs text-slate-500 mt-1.5 flex items-center gap-1">
                <MessageSquare className="w-3 h-3" />
                {review.contract.title}
              </div>
            )}
            {/* Score breakdown */}
            <div className="flex items-center gap-4 mt-2.5">
              {[
                { label: 'Quality', value: review.quality },
                { label: 'Communication', value: review.communication },
                { label: 'Professionalism', value: review.professionalism },
              ].map(score => (
                <div key={score.label} className="text-center">
                  <div className="text-sm font-semibold text-slate-900">{score.value.toFixed(1)}</div>
                  <div className="text-xs text-slate-400">{score.label}</div>
                </div>
              ))}
            </div>
            {review.comment && (
              <p className="text-sm text-slate-600 mt-3 leading-relaxed">{review.comment}</p>
            )}
          </div>
        </div>
      </CardContent>
    </Card>
  )

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-32" />
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
        </div>
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <Skeleton className="h-64 rounded-xl" />
          <Skeleton className="lg:col-span-2 h-64 rounded-xl" />
        </div>
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Reviews</h1>
        <p className="text-slate-500 mt-1">View and manage your reviews.</p>
      </div>

      {/* Summary Stats */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Total Reviews"
          value={stats.total}
          icon={Star}
          iconColor="bg-amber-100 text-amber-600"
        />
        <StatCard
          label="Average Rating"
          value={stats.averageRating}
          suffix="/5"
          icon={Star}
          iconColor="bg-emerald-100 text-emerald-600"
          decimals={1}
        />
        <StatCard
          label="5-Star Reviews"
          value={reviews.filter(r => Math.floor(r.rating) === 5).length}
          icon={Star}
          iconColor="bg-emerald-100 text-emerald-600"
        />
        <StatCard
          label="4-Star Reviews"
          value={reviews.filter(r => Math.floor(r.rating) === 4).length}
          icon={Star}
          iconColor="bg-amber-100 text-amber-600"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left: Stats */}
        <Card className="border-slate-200 h-fit">
          <CardContent className="p-6">
            <div className="text-center mb-6">
              <div className="text-4xl font-bold text-slate-900">{stats.averageRating.toFixed(1)}</div>
              <StarRating rating={Math.round(stats.averageRating)} />
              <div className="text-sm text-slate-500 mt-1">{stats.total} reviews</div>
            </div>
            <div className="space-y-2.5">
              {distribution.map(d => (
                <div key={d.star} className="flex items-center gap-2">
                  <span className="text-xs text-slate-500 w-6">{d.star}★</span>
                  <Progress value={d.percentage} className="h-2 flex-1" />
                  <span className="text-xs text-slate-500 w-8 text-right">{d.count}</span>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>

        {/* Right: Reviews */}
        <div className="lg:col-span-2">
          <Tabs defaultValue="received">
            <TabsList className="bg-slate-100 mb-4">
              <TabsTrigger value="received" className="text-sm">Received ({reviews.length})</TabsTrigger>
              <TabsTrigger value="given" className="text-sm">Given (0)</TabsTrigger>
            </TabsList>

            <TabsContent value="received">
              {reviews.length === 0 ? (
                <Card className="border-slate-200">
                  <CardContent className="p-12">
                    <EmptyState
                      icon={Star}
                      title="No reviews yet"
                      description="Reviews will appear here after you complete contracts and receive feedback."
                    />
                  </CardContent>
                </Card>
              ) : (
                <div className="space-y-3 max-h-[600px] overflow-y-auto">
                  {reviews.map(review => (
                    <ReviewCard key={review.id} review={review} />
                  ))}
                </div>
              )}
            </TabsContent>

            <TabsContent value="given">
              <Card className="border-slate-200">
                <CardContent className="p-12">
                  <EmptyState
                    icon={User}
                    title="No reviews given"
                    description="After completing a contract, you'll be able to leave a review for the other party."
                  />
                </CardContent>
              </Card>
            </TabsContent>
          </Tabs>
        </div>
      </div>
    </motion.div>
  )
}
