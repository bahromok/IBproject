'use client'

import { useState, useEffect } from 'react'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Button } from '@/components/ui/button'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import {
  Search,
  Bell,
  MessageSquare,
  Menu,
  Settings,
  User,
  LogOut,
  ChevronDown,
  Briefcase,
  X,
  CheckCheck,
  ArrowRightLeft,
} from 'lucide-react'
import { timeAgo } from '@/lib/format'

interface NotificationItem {
  id: string
  type: string
  title: string
  message: string
  isRead: boolean
  actionUrl?: string | null
  createdAt: string
}

interface TopNavProps {
  onToggleSidebar: () => void
}

export function TopNav({ onToggleSidebar }: TopNavProps) {
  const [searchQuery, setSearchQuery] = useState('')
  const [isSearchOpen, setIsSearchOpen] = useState(false)
  const { user, logout } = useAuthStore()
  const { navigate } = useNavigationStore()
  const { toast } = useToast()

  // Notification state
  const [notifications, setNotifications] = useState<NotificationItem[]>([])
  const [unreadCount, setUnreadCount] = useState(0)
  const [isNotifOpen, setIsNotifOpen] = useState(false)

  // ─── Poll notifications on dropdown open & every 30s ─────────────
  const refreshNotifications = async () => {
    const u = useAuthStore.getState().user
    if (!u?.id) return
    try {
      const [notifRes, unreadRes] = await Promise.all([
        authFetch('/api/notifications?limit=5'),
        authFetch('/api/notifications?unreadOnly=true&limit=1'),
      ])
      if (notifRes.ok) {
        const data = await notifRes.json()
        setNotifications(data.data?.notifications || [])
      }
      if (unreadRes.ok) {
        const data = await unreadRes.json()
        setUnreadCount(data.data?.unreadCount || 0)
      }
    } catch {
      // silently fail
    }
  }

  useEffect(() => {
    const id = setInterval(refreshNotifications, 30000)
    const t = setTimeout(refreshNotifications, 100)
    return () => { clearInterval(id); clearTimeout(t) }
  }, [])

  // ─── Mark notification as read ───────────────────────────────────
  const markAsRead = async (notification: NotificationItem) => {
    if (!notification.isRead) {
      authFetch('/api/notifications', {
        method: 'PATCH',
        body: JSON.stringify({ notificationId: notification.id }),
      }).catch(() => { })
      setNotifications((prev) =>
        prev.map((n) => (n.id === notification.id ? { ...n, isRead: true } : n))
      )
      setUnreadCount((prev) => Math.max(0, prev - 1))
    }
    if (notification.actionUrl) {
      navigateToUrl(notification.actionUrl)
    }
  }

  // ─── Mark all as read ────────────────────────────────────────────
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

  // ─── Parse actionUrl to navigate ─────────────────────────────────
  const navigateToUrl = (url: string) => {
    const urlObj = new URL(url, 'http://localhost')
    const pathname = urlObj.pathname.replace(/^\//, '')
    const params: Record<string, string> = {}
    urlObj.searchParams.forEach((v, k) => {
      params[k] = v
    })
    navigate(pathname as any, params)
  }

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault()
    if (searchQuery.trim()) {
      navigate('jobs', { search: searchQuery.trim() })
      setSearchQuery('')
      setIsSearchOpen(false)
    }
  }

  const handleLogout = async () => {
    await logout()
    toast({ title: 'Signed out', description: 'You have been signed out successfully.' })
    navigate('landing')
  }

  const initials = user?.name
    ? user.name.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2)
    : user?.email?.[0]?.toUpperCase() || 'U'

  const roleIcon = user?.role === 'ADMIN' ? (
    <Badge className="bg-amber-100 text-amber-700 border-amber-200 text-xs px-1.5 py-0">Admin</Badge>
  ) : user?.role === 'CLIENT' ? (
    <Badge className="bg-emerald-100 text-emerald-700 border-emerald-200 text-xs px-1.5 py-0">Client</Badge>
  ) : null

  return (
    <header className="sticky top-0 z-30 h-16 border-b border-slate-200 bg-white/80 backdrop-blur-md">
      <div className="flex items-center justify-between h-full px-4 lg:px-6">
        {/* Left: Hamburger + Search */}
        <div className="flex items-center gap-3 flex-1">
          {/* Mobile hamburger */}
          <Button
            variant="ghost"
            size="icon"
            className="lg:hidden text-slate-600 hover:text-slate-900 h-9 w-9"
            onClick={onToggleSidebar}
            aria-label="Toggle sidebar"
          >
            <Menu className="w-5 h-5" />
          </Button>

          {/* Search Bar */}
          <form onSubmit={handleSearch} className="hidden sm:flex items-center flex-1 max-w-md">
            <div className="relative w-full">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <Input
                type="text"
                placeholder="Search jobs, freelancers, projects..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="h-9 pl-9 pr-4 border-slate-200 bg-slate-50/50 focus:bg-white focus:border-emerald-400 focus:ring-emerald-500/20 text-sm rounded-lg"
              />
            </div>
          </form>

          {/* Mobile Search Button */}
          <Button
            variant="ghost"
            size="icon"
            className="sm:hidden text-slate-600 hover:text-slate-900 h-9 w-9"
            onClick={() => setIsSearchOpen(!isSearchOpen)}
            aria-label="Search"
          >
            {isSearchOpen ? <X className="w-5 h-5" /> : <Search className="w-5 h-5" />}
          </Button>
        </div>

        {/* Mobile search bar */}
        {isSearchOpen && (
          <form onSubmit={handleSearch} className="sm:hidden absolute top-16 left-0 right-0 p-3 bg-white border-b z-20">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <Input
                autoFocus
                type="text"
                placeholder="Search jobs, freelancers..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="h-10 pl-9 pr-4 border-slate-200 bg-slate-50 text-sm rounded-lg"
              />
            </div>
          </form>
        )}

        {/* Right: Actions */}
        <div className="flex items-center gap-1.5">
          {/* Post Job Button (Client only) */}
          {user?.role === 'CLIENT' && (
            <Button
              onClick={() => navigate('jobs/post')}
              className="hidden sm:flex bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg h-8 text-xs font-medium px-3 shadow-sm"
            >
              <Briefcase className="w-3.5 h-3.5 mr-1.5" />
              Post a Job
            </Button>
          )}

          {/* Notifications Dropdown */}
          <DropdownMenu open={isNotifOpen} onOpenChange={setIsNotifOpen}>
            <DropdownMenuTrigger asChild>
              <Button
                variant="ghost"
                className="relative text-slate-600 hover:text-slate-900 h-9 w-9"
                aria-label="Notifications"
              >
                <Bell className="w-5 h-5" />
                {unreadCount > 0 && (
                  <span className="absolute -top-0.5 -right-0.5 min-w-[18px] h-[18px] flex items-center justify-center bg-emerald-500 text-white text-[10px] font-bold rounded-full px-1">
                    {unreadCount > 99 ? '99+' : unreadCount}
                  </span>
                )}
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-80 p-0">
              <DropdownMenuLabel className="flex items-center justify-between px-4 py-3">
                <span className="font-semibold text-slate-900">Notifications</span>
                {unreadCount > 0 && (
                  <Badge variant="secondary" className="bg-emerald-100 text-emerald-700 text-xs font-medium">
                    {unreadCount} new
                  </Badge>
                )}
              </DropdownMenuLabel>
              <DropdownMenuSeparator />

              {notifications.length === 0 ? (
                <div className="py-8 text-center">
                  <Bell className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                  <p className="text-sm text-slate-500">No notifications yet</p>
                  <p className="text-xs text-slate-400 mt-1">We&apos;ll let you know when something arrives</p>
                </div>
              ) : (
                <>
                  <div className="max-h-80 overflow-y-auto">
                    {notifications.map((notif) => (
                      <DropdownMenuItem
                        key={notif.id}
                        className={`flex items-start gap-3 py-3 px-4 cursor-pointer border-b border-slate-50 last:border-0 ${!notif.isRead ? 'bg-emerald-50/50' : ''
                          }`}
                        onClick={() => markAsRead(notif)}
                      >
                        <div className="flex-shrink-0 mt-0.5">
                          <NotifIcon type={notif.type} />
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-start justify-between gap-2">
                            <p className={`text-sm leading-tight ${!notif.isRead ? 'font-semibold text-slate-900' : 'font-medium text-slate-700'}`}>
                              {notif.title}
                            </p>
                            {!notif.isRead && (
                              <div className="w-2 h-2 bg-emerald-500 rounded-full flex-shrink-0 mt-1.5" />
                            )}
                          </div>
                          <p className="text-xs text-slate-500 mt-1 line-clamp-2">{notif.message}</p>
                          <p className="text-[11px] text-slate-400 mt-1">{timeAgo(notif.createdAt)}</p>
                        </div>
                      </DropdownMenuItem>
                    ))}
                  </div>

                  <DropdownMenuSeparator />
                  <div className="flex items-center">
                    {unreadCount > 0 && (
                      <DropdownMenuItem
                        className="flex-1 text-center justify-center text-xs text-emerald-600 font-medium py-2.5 cursor-pointer hover:bg-emerald-50"
                        onClick={(e) => {
                          e.stopPropagation()
                          markAllAsRead()
                        }}
                      >
                        <CheckCheck className="w-3.5 h-3.5 mr-1.5" />
                        Mark all as read
                      </DropdownMenuItem>
                    )}
                    <DropdownMenuItem
                      className="flex-1 text-center justify-center text-xs text-slate-600 font-medium py-2.5 cursor-pointer"
                      onClick={() => {
                        setIsNotifOpen(false)
                        navigate('notifications')
                      }}
                    >
                      View all notifications
                    </DropdownMenuItem>
                  </div>
                </>
              )}
            </DropdownMenuContent>
          </DropdownMenu>

          {/* Messages */}
          <Button
            variant="ghost"
            className="relative text-slate-600 hover:text-slate-900 h-9 w-9"
            onClick={() => navigate('messages')}
            aria-label="Messages"
          >
            <MessageSquare className="w-5 h-5" />
          </Button>

          {/* User Menu */}
          <DropdownMenu>
            <DropdownMenuTrigger asChild>
              <Button
                variant="ghost"
                className="flex items-center gap-2 h-9 px-2 hover:bg-slate-100 rounded-lg"
              >
                <Avatar className="h-7 w-7 border border-slate-200">
                  <AvatarFallback className="bg-emerald-100 text-emerald-700 text-xs font-semibold">
                    {initials}
                  </AvatarFallback>
                </Avatar>
                <span className="hidden md:block text-sm font-medium text-slate-700 max-w-[120px] truncate">
                  {user?.name || user?.email || 'User'}
                </span>
                <ChevronDown className="hidden md:block w-3.5 h-3.5 text-slate-400" />
              </Button>
            </DropdownMenuTrigger>
            <DropdownMenuContent align="end" className="w-56">
              <DropdownMenuLabel className="flex items-center gap-2 py-2">
                <div className="flex-1 min-w-0">
                  <div className="text-sm font-medium text-slate-900 truncate">
                    {user?.name || user?.email || 'User'}
                  </div>
                  <div className="text-xs text-slate-500 truncate">{user?.email}</div>
                </div>
                {roleIcon}
              </DropdownMenuLabel>
              <DropdownMenuSeparator />
              <DropdownMenuItem
                onClick={() => navigate('profile')}
                className="cursor-pointer"
              >
                <User className="w-4 h-4 mr-2 text-slate-400" />
                Profile
              </DropdownMenuItem>
              <DropdownMenuItem
                onClick={async () => {
                  const currentRole = user?.role
                  await useAuthStore.getState().switchRole()
                  navigate('dashboard')
                  toast({
                    title: 'Role Switched',
                    description: `Switched to ${currentRole === 'CLIENT' ? 'Freelancer' : 'Client'} mode.`
                  })
                }}
                className="cursor-pointer font-medium text-emerald-600 focus:text-emerald-700 focus:bg-emerald-50"
              >
                <ArrowRightLeft className="w-4 h-4 mr-2" />
                Switch to {user?.role === 'CLIENT' ? 'Freelancer' : 'Client'}
              </DropdownMenuItem>
              <DropdownMenuItem
                onClick={() => navigate('settings')}
                className="cursor-pointer"
              >
                <Settings className="w-4 h-4 mr-2 text-slate-400" />
                Settings
              </DropdownMenuItem>
              <DropdownMenuSeparator />
              <DropdownMenuItem
                onClick={handleLogout}
                className="cursor-pointer text-red-600 focus:text-red-600"
              >
                <LogOut className="w-4 h-4 mr-2" />
                Sign out
              </DropdownMenuItem>
            </DropdownMenuContent>
          </DropdownMenu>
        </div>
      </div>
    </header>
  )
}

// ─── Notification type icon helper ──────────────────────────────────

function NotifIcon({ type }: { type: string }) {
  const baseClass = 'w-8 h-8 rounded-full flex items-center justify-center text-xs'

  switch (type) {
    case 'NEW_PROPOSAL':
      return <div className={`${baseClass} bg-blue-100 text-blue-600`}><Briefcase className="w-4 h-4" /></div>
    case 'PROPOSAL_ACCEPTED':
      return <div className={`${baseClass} bg-emerald-100 text-emerald-600`}><CheckCheck className="w-4 h-4" /></div>
    case 'PROPOSAL_REJECTED':
      return <div className={`${baseClass} bg-red-100 text-red-600`}><X className="w-4 h-4" /></div>
    case 'NEW_MESSAGE':
      return <div className={`${baseClass} bg-violet-100 text-violet-600`}><MessageSquare className="w-4 h-4" /></div>
    case 'MILESTONE_APPROVED':
      return <div className={`${baseClass} bg-emerald-100 text-emerald-600`}><CheckCheck className="w-4 h-4" /></div>
    case 'PAYMENT_RECEIVED':
      return <div className={`${baseClass} bg-amber-100 text-amber-600`}><Briefcase className="w-4 h-4" /></div>
    case 'CONTRACT_COMPLETED':
      return <div className={`${baseClass} bg-emerald-100 text-emerald-600`}><CheckCheck className="w-4 h-4" /></div>
    case 'NEW_REVIEW':
      return <div className={`${baseClass} bg-yellow-100 text-yellow-600`}><User className="w-4 h-4" /></div>
    case 'REPORT_STATUS':
      return <div className={`${baseClass} bg-slate-100 text-slate-600`}><Settings className="w-4 h-4" /></div>
    default:
      return <div className={`${baseClass} bg-slate-100 text-slate-600`}><Bell className="w-4 h-4" /></div>
  }
}
