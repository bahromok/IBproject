'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { authFetch } from '@/lib/api-fetch'
import { useNavigationStore } from '@/store/navigation-store'
import { useAuthStore } from '@/store/auth-store'
import { timeAgo } from '@/lib/format'
import { EmptyState } from '@/components/shared/empty-state'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import {
  Bell,
  CheckCheck,
  Briefcase,
  MessageSquare,
  User,
  Settings,
  X,
  Filter,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react'
import { Skeleton } from '@/components/ui/skeleton'

// ─── Types ──────────────────────────────────────────────────────────

interface NotificationItem {
  id: string
  type: string
  title: string
  message: string
  isRead: boolean
  actionUrl?: string | null
  createdAt: string
}

type FilterTab = 'all' | 'unread' | 'read'

// ─── Notification type config ───────────────────────────────────────

const NOTIF_TYPE_CONFIG: Record<string, { icon: typeof Bell; color: string; label: string }> = {
  NEW_PROPOSAL: { icon: Briefcase, color: 'bg-blue-100 text-blue-600', label: 'Proposal' },
  PROPOSAL_ACCEPTED: { icon: CheckCheck, color: 'bg-emerald-100 text-emerald-600', label: 'Accepted' },
  PROPOSAL_REJECTED: { icon: X, color: 'bg-red-100 text-red-600', label: 'Rejected' },
  NEW_MESSAGE: { icon: MessageSquare, color: 'bg-violet-100 text-violet-600', label: 'Message' },
  MILESTONE_APPROVED: { icon: CheckCheck, color: 'bg-emerald-100 text-emerald-600', label: 'Milestone' },
  PAYMENT_RECEIVED: { icon: Briefcase, color: 'bg-amber-100 text-amber-600', label: 'Payment' },
  CONTRACT_COMPLETED: { icon: CheckCheck, color: 'bg-emerald-100 text-emerald-600', label: 'Contract' },
  NEW_REVIEW: { icon: User, color: 'bg-yellow-100 text-yellow-600', label: 'Review' },
  REPORT_STATUS: { icon: Settings, color: 'bg-slate-100 text-slate-600', label: 'Report' },
  SYSTEM: { icon: Bell, color: 'bg-slate-100 text-slate-600', label: 'System' },
}

// ─── Component ──────────────────────────────────────────────────────

export function NotificationsPage() {
  const { navigate } = useNavigationStore()
  const { isAuthenticated } = useAuthStore()

  const [notifications, setNotifications] = useState<NotificationItem[]>([])
  const [unreadCount, setUnreadCount] = useState(0)
  const [loading, setLoading] = useState(true)
  const [filter, setFilter] = useState<FilterTab>('all')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)

  const fetchNotifications = useCallback(async () => {
    if (!isAuthenticated) return
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', String(page))
      params.set('limit', '15')
      if (filter === 'unread') params.set('unreadOnly', 'true')
      if (filter === 'read') params.set('unreadOnly', 'false')

      const res = await authFetch(`/api/notifications?${params.toString()}`)
      if (res.ok) {
        const data = await res.json()
        setNotifications(data.data?.notifications || [])
        setUnreadCount(data.data?.unreadCount || 0)
        setTotalPages(data.data?.pagination?.totalPages || 1)
      }
    } catch {
      // silently fail
    } finally {
      setLoading(false)
    }
  }, [isAuthenticated, filter, page])

  useEffect(() => {
    fetchNotifications()
  }, [fetchNotifications])

  // Reset to page 1 on filter change
  useEffect(() => {
    setPage(1)
  }, [filter])

  // ─── Mark as read ───────────────────────────────────────────────
  const markAsRead = async (notification: NotificationItem) => {
    if (!notification.isRead) {
      authFetch('/api/notifications', {
        method: 'PATCH',
        body: JSON.stringify({ notificationId: notification.id }),
      }).catch(() => {})
      setNotifications((prev) =>
        prev.map((n) => (n.id === notification.id ? { ...n, isRead: true } : n))
      )
      setUnreadCount((prev) => Math.max(0, prev - 1))
    }
    if (notification.actionUrl) {
      navigateToUrl(notification.actionUrl)
    }
  }

  // ─── Mark all as read ──────────────────────────────────────────
  const markAllAsRead = async () => {
    try {
      await authFetch('/api/notifications', {
        method: 'PATCH',
        body: JSON.stringify({ markAll: true }),
      })
      setNotifications((prev) => prev.map((n) => ({ ...n, isRead: true })))
      setUnreadCount(0)
    } catch {
      // silently fail
    }
  }

  // ─── Parse actionUrl ───────────────────────────────────────────
  const navigateToUrl = (url: string) => {
    try {
      const urlObj = new URL(url, 'http://localhost')
      const pathname = urlObj.pathname.replace(/^\//, '')
      const params: Record<string, string> = {}
      urlObj.searchParams.forEach((v, k) => {
        params[k] = v
      })
      navigate(pathname as any, params)
    } catch {
      // ignore invalid URLs
    }
  }

  // ─── Filtered notifications ────────────────────────────────────
  const filteredNotifications = notifications

  // ─── Render ────────────────────────────────────────────────────
  return (
    <motion.div
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      className="max-w-4xl mx-auto"
    >
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mb-6">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Notifications</h1>
          <p className="text-sm text-slate-500 mt-1">
            {unreadCount > 0
              ? `You have ${unreadCount} unread notification${unreadCount !== 1 ? 's' : ''}`
              : 'All caught up!'}
          </p>
        </div>
        {unreadCount > 0 && (
          <Button
            variant="outline"
            size="sm"
            className="border-emerald-200 text-emerald-700 hover:bg-emerald-50"
            onClick={markAllAsRead}
          >
            <CheckCheck className="w-4 h-4 mr-1.5" />
            Mark all as read
          </Button>
        )}
      </div>

      {/* Filter tabs */}
      <div className="flex items-center gap-2 mb-4">
        {(['all', 'unread', 'read'] as FilterTab[]).map((tab) => (
          <button
            key={tab}
            onClick={() => setFilter(tab)}
            className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors ${
              filter === tab
                ? 'bg-emerald-100 text-emerald-700'
                : 'text-slate-600 hover:bg-slate-100'
            }`}
          >
            {tab === 'all' && 'All'}
            {tab === 'unread' && 'Unread'}
            {tab === 'read' && 'Read'}
          </button>
        ))}
      </div>

      {/* Content */}
      {loading ? (
        <div className="space-y-3">
          {Array.from({ length: 5 }).map((_, i) => (
            <div key={i} className="flex items-start gap-4 p-4 rounded-xl border border-slate-100 bg-white">
              <Skeleton className="w-10 h-10 rounded-full flex-shrink-0" />
              <div className="flex-1 space-y-2">
                <Skeleton className="h-4 w-3/4" />
                <Skeleton className="h-3 w-full" />
                <Skeleton className="h-3 w-20" />
              </div>
            </div>
          ))}
        </div>
      ) : filteredNotifications.length === 0 ? (
        <EmptyState
          icon={Bell}
          title={filter === 'unread' ? 'No unread notifications' : 'No notifications'}
          description={
            filter === 'unread'
              ? 'You have read all your notifications. Keep up the great work!'
              : 'When you receive notifications about proposals, messages, payments, and more, they will appear here.'
          }
        />
      ) : (
        <>
          <div className="space-y-2">
            {filteredNotifications.map((notif) => {
              const config = NOTIF_TYPE_CONFIG[notif.type] || NOTIF_TYPE_CONFIG.SYSTEM
              const Icon = config.icon
              return (
                <motion.div
                  key={notif.id}
                  initial={{ opacity: 0, x: -5 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={{ duration: 0.15 }}
                  onClick={() => markAsRead(notif)}
                  className={`flex items-start gap-4 p-4 rounded-xl border cursor-pointer transition-all hover:shadow-sm ${
                    notif.isRead
                      ? 'border-slate-100 bg-white'
                      : 'border-emerald-100 bg-emerald-50/40'
                  }`}
                >
                  {/* Icon */}
                  <div className={`w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ${config.color}`}>
                    <Icon className="w-5 h-5" />
                  </div>

                  {/* Content */}
                  <div className="flex-1 min-w-0">
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <p className={`text-sm leading-tight ${notif.isRead ? 'font-medium text-slate-700' : 'font-semibold text-slate-900'}`}>
                            {notif.title}
                          </p>
                          <Badge variant="secondary" className="text-[10px] px-1.5 py-0 bg-slate-100 text-slate-500 font-normal hidden sm:inline-flex">
                            {config.label}
                          </Badge>
                        </div>
                        <p className="text-sm text-slate-500 mt-1 line-clamp-2">{notif.message}</p>
                      </div>
                      {!notif.isRead && (
                        <div className="w-2.5 h-2.5 bg-emerald-500 rounded-full flex-shrink-0 mt-1.5" />
                      )}
                    </div>
                    <p className="text-xs text-slate-400 mt-2">{timeAgo(notif.createdAt)}</p>
                  </div>
                </motion.div>
              )
            })}
          </div>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-center gap-2 mt-6">
              <Button
                variant="outline"
                size="sm"
                disabled={page <= 1}
                onClick={() => setPage((p) => Math.max(1, p - 1))}
                className="border-slate-200"
              >
                <ChevronLeft className="w-4 h-4" />
                Previous
              </Button>
              <span className="text-sm text-slate-500">
                Page {page} of {totalPages}
              </span>
              <Button
                variant="outline"
                size="sm"
                disabled={page >= totalPages}
                onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
                className="border-slate-200"
              >
                Next
                <ChevronRight className="w-4 h-4" />
              </Button>
            </div>
          )}
        </>
      )}
    </motion.div>
  )
}
