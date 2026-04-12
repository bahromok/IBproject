'use client'

import { useEffect, useState } from 'react'
import { Avatar, AvatarFallback, AvatarImage } from '@/components/ui/avatar'
import { Badge } from '@/components/ui/badge'
import { Star } from 'lucide-react'
import { cn } from '@/lib/utils'
import { authFetch } from '@/lib/api-fetch'

interface IdentityProtectedUserProps {
  userId: string
  size?: 'sm' | 'md' | 'lg'
  showRating?: boolean
  showRole?: boolean
  className?: string
}

interface UserData {
  displayName: string
  avatar: string | null
  role: string
  rating: number
  totalReviews: number
}

export function IdentityProtectedUser({
  userId,
  size = 'md',
  showRating = false,
  showRole = false,
  className,
}: IdentityProtectedUserProps) {
  const [user, setUser] = useState<UserData | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    const fetchUser = async () => {
      try {
        const res = await authFetch(`/api/freelancers/${userId}`)
        if (res.ok) {
          const data = await res.json()
          if (!cancelled) {
            setUser({
              displayName: data.data?.displayName || data.displayName || 'Anonymous',
              avatar: data.data?.avatar || data.avatar || null,
              role: data.data?.role || 'FREELANCER',
              rating: data.data?.rating || data.rating || 0,
              totalReviews: data.data?.totalReviews || data.totalReviews || 0,
            })
          }
        } else {
          // Try as client profile fallback
          const res2 = await authFetch(`/api/conversations`)
          if (res2.ok) {
            const data2 = await res2.json()
            const conv = (data2.conversations || []).find(
              (c: any) => c.otherUser?.id === userId
            )
            if (conv?.otherUser && !cancelled) {
              setUser({
                displayName: conv.otherUser.displayName,
                avatar: conv.otherUser.avatar,
                role: 'CLIENT',
                rating: 0,
                totalReviews: 0,
              })
            }
          }
        }
      } catch {
        // Silently fail - this is for display only
      } finally {
        if (!cancelled) setLoading(false)
      }
    }
    fetchUser()
    return () => { cancelled = true }
  }, [userId])

  const sizeClasses = {
    sm: { avatar: 'h-7 w-7 text-xs', name: 'text-xs font-medium', star: 'w-3 h-3' },
    md: { avatar: 'h-9 w-9 text-sm', name: 'text-sm font-medium', star: 'w-3.5 h-3.5' },
    lg: { avatar: 'h-12 w-12 text-base', name: 'text-base font-semibold', star: 'w-4 h-4' },
  }

  const s = sizeClasses[size]

  if (loading) {
    return (
      <div className={cn('flex items-center gap-2', className)}>
        <div className={cn('rounded-full bg-slate-200 animate-pulse', s.avatar)} />
        <div className="h-3 w-20 bg-slate-200 rounded animate-pulse" />
      </div>
    )
  }

  if (!user) {
    return (
      <div className={cn('flex items-center gap-2', className)}>
        <Avatar className={s.avatar}>
          <AvatarFallback className="bg-slate-100 text-slate-500">?</AvatarFallback>
        </Avatar>
        <span className={cn('text-slate-500', s.name)}>Anonymous</span>
      </div>
    )
  }

  const initials = user.displayName.slice(0, 2).toUpperCase()

  return (
    <div className={cn('flex items-center gap-2', className)}>
      <Avatar className={s.avatar}>
        {user.avatar && <AvatarImage src={user.avatar} alt={user.displayName} />}
        <AvatarFallback className="bg-emerald-100 text-emerald-700 font-semibold">
          {initials}
        </AvatarFallback>
      </Avatar>
      <div className="flex flex-col">
        <div className="flex items-center gap-1.5">
          <span className={cn('text-slate-900', s.name)}>{user.displayName}</span>
          {showRole && (
            <Badge
              variant="secondary"
              className={cn(
                'px-1.5 py-0 text-[10px] font-medium',
                user.role === 'CLIENT'
                  ? 'bg-amber-100 text-amber-700'
                  : user.role === 'ADMIN'
                    ? 'bg-red-100 text-red-700'
                    : 'bg-emerald-100 text-emerald-700'
              )}
            >
              {user.role === 'CLIENT' ? 'Client' : user.role === 'ADMIN' ? 'Admin' : 'Freelancer'}
            </Badge>
          )}
        </div>
        {showRating && user.rating > 0 && (
          <div className="flex items-center gap-1">
            <Star className={cn('fill-amber-400 text-amber-400', s.star)} />
            <span className="text-xs text-slate-500">
              {user.rating.toFixed(1)} ({user.totalReviews})
            </span>
          </div>
        )}
      </div>
    </div>
  )
}
