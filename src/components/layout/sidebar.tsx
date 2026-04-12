'use client'

import { motion, AnimatePresence } from 'framer-motion'
import { useNavigationStore, type PageName } from '@/store/navigation-store'
import { useAuthStore } from '@/store/auth-store'
import { ScrollArea } from '@/components/ui/scroll-area'
import { Badge } from '@/components/ui/badge'
import { Separator } from '@/components/ui/separator'
import { cn } from '@/lib/utils'
import {
  LayoutDashboard,
  Search,
  Briefcase,
  PlusCircle,
  FileText,
  MessageSquare,
  User,
  Star,
  Settings,
  AlertTriangle,
  Bookmark,
  Users,
  Shield,
  BarChart3,
  FileWarning,
  Scale,
  GraduationCap,
  X,
  ChevronLeft,
  ShieldAlert,
  Bell,
} from 'lucide-react'

interface NavItem {
  label: string
  icon: React.ElementType
  view: PageName
  badge?: string | number
  badgeVariant?: 'default' | 'secondary' | 'destructive'
}

interface NavSection {
  title?: string
  items: NavItem[]
}

function getClientNav(): NavSection[] {
  return [
    {
      items: [
        { label: 'Dashboard', icon: LayoutDashboard, view: 'dashboard' },
        { label: 'Find Freelancers', icon: Users, view: 'freelancers' },
        { label: 'My Jobs', icon: Briefcase, view: 'jobs/my-jobs' },
        { label: 'Post a Job', icon: PlusCircle, view: 'jobs/post' },
      ],
    },
    {
      items: [
        { label: 'Contracts', icon: FileText, view: 'contracts' },
        { label: 'Messages', icon: MessageSquare, view: 'messages' },
        { label: 'Notifications', icon: Bell, view: 'notifications' },
        { label: 'Saved Jobs', icon: Bookmark, view: 'saved-jobs' },
      ],
    },
    {
      items: [
        { label: 'Profile', icon: User, view: 'profile' },
        { label: 'Settings', icon: Settings, view: 'settings' },
        { label: 'File a Report', icon: AlertTriangle, view: 'reports' },
      ],
    },
  ]
}

function getFreelancerNav(): NavSection[] {
  return [
    {
      items: [
        { label: 'Dashboard', icon: LayoutDashboard, view: 'dashboard' },
        { label: 'Find Jobs', icon: Search, view: 'jobs' },
        { label: 'My Proposals', icon: FileText, view: 'proposals' },
      ],
    },
    {
      items: [
        { label: 'Contracts', icon: Briefcase, view: 'contracts' },
        { label: 'Messages', icon: MessageSquare, view: 'messages' },
        { label: 'Notifications', icon: Bell, view: 'notifications' },
      ],
    },
    {
      items: [
        { label: 'Profile', icon: User, view: 'profile' },
        { label: 'Reviews', icon: Star, view: 'reviews' },
        { label: 'Settings', icon: Settings, view: 'settings' },
      ],
    },
  ]
}

function getAdminNav(): NavSection[] {
  return [
    {
      items: [
        { label: 'Dashboard', icon: LayoutDashboard, view: 'admin/dashboard' },
        { label: 'Users', icon: Users, view: 'admin/users' },
        { label: 'Jobs Moderation', icon: Briefcase, view: 'admin/jobs' },
      ],
    },
    {
      items: [
        { label: 'Reports', icon: AlertTriangle, view: 'admin/reports', badge: 5, badgeVariant: 'destructive' },
        { label: 'Fraud Alerts', icon: ShieldAlert, view: 'admin/fraud', badge: 0, badgeVariant: 'destructive' },
        { label: 'Disputes', icon: Scale, view: 'admin/disputes', badge: 2, badgeVariant: 'destructive' },
      ],
    },
    {
      items: [
        { label: 'Analytics', icon: BarChart3, view: 'admin/analytics' },
      ],
    },
  ]
}

interface SidebarProps {
  isOpen: boolean
  onClose: () => void
  isCollapsed: boolean
  onToggleCollapse: () => void
}

export function Sidebar({ isOpen, onClose, isCollapsed, onToggleCollapse }: SidebarProps) {
  const { currentPage, navigate } = useNavigationStore()
  const { user } = useAuthStore()

  const role = user?.role || 'FREELANCER'
  const sections = role === 'ADMIN' ? getAdminNav() : role === 'CLIENT' ? getClientNav() : getFreelancerNav()

  const handleNavClick = (view: PageName) => {
    navigate(view)
    onClose()
  }

  const isActive = (view: PageName) => {
    if (view === 'dashboard' && currentPage === 'dashboard') return true
    if (view === 'admin/dashboard' && currentPage.startsWith('admin')) return true
    if (view === 'contracts' && (currentPage === 'contracts' || currentPage === 'contracts/detail')) return true
    if (view === 'jobs' && currentPage.startsWith('jobs') && currentPage !== 'jobs/post') return true
    if (view === 'messages' && (currentPage === 'messages' || currentPage === 'messages/conversation')) return true
    if (view === 'jobs/post' && currentPage === 'jobs/post') return true
    return currentPage === view
  }

  const sidebarContent = (
    <div className="flex flex-col h-full">
      {/* Logo */}
      <div className={cn(
        'flex items-center h-16 px-4 border-b border-slate-200/80 shrink-0',
        isCollapsed ? 'justify-center' : 'gap-2'
      )}>
        {!isCollapsed && (
          <button onClick={() => { navigate('dashboard'); onClose() }} className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center shrink-0">
              <GraduationCap className="w-4.5 h-4.5 text-white" />
            </div>
            <span className="text-lg font-bold text-slate-900">
              Student<span className="text-emerald-500">Hire</span>
            </span>
          </button>
        )}
        {isCollapsed && (
          <button onClick={() => { navigate('dashboard'); onClose() }} className="flex items-center justify-center">
            <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center">
              <GraduationCap className="w-4.5 h-4.5 text-white" />
            </div>
          </button>
        )}
      </div>

      {/* Role Badge */}
      {!isCollapsed && (
        <div className="px-4 pt-4 pb-2">
          <Badge
            variant="secondary"
            className={cn(
              'w-full justify-center py-1 text-xs font-medium',
              role === 'ADMIN'
                ? 'bg-amber-100 text-amber-700 border-amber-200'
                : role === 'CLIENT'
                  ? 'bg-emerald-100 text-emerald-700 border-emerald-200'
                  : 'bg-slate-100 text-slate-600 border-slate-200'
            )}
          >
            {role === 'ADMIN' && <Shield className="w-3 h-3 mr-1" />}
            {role === 'CLIENT' ? 'Client Account' : role === 'FREELANCER' ? 'Freelancer Account' : 'Admin Panel'}
          </Badge>
        </div>
      )}

      {/* Navigation Items */}
      <ScrollArea className="flex-1 px-3 py-2">
        <nav className="space-y-1">
          {sections.map((section, sectionIdx) => (
            <div key={sectionIdx}>
              {section.title && !isCollapsed && (
                <>
                  <div className="px-3 py-2 text-xs font-semibold text-slate-400 uppercase tracking-wider">
                    {section.title}
                  </div>
                </>
              )}
              {sectionIdx > 0 && <Separator className="my-2 bg-slate-200/50" />}
              <div className="space-y-0.5">
                {section.items.map((item) => {
                  const active = isActive(item.view)
                  return (
                    <button
                      key={item.view}
                      onClick={() => handleNavClick(item.view)}
                      className={cn(
                        'w-full flex items-center gap-3 rounded-lg text-sm font-medium transition-all duration-200',
                        isCollapsed ? 'justify-center p-2.5' : 'px-3 py-2.5',
                        active
                          ? 'bg-emerald-50 text-emerald-700 shadow-sm'
                          : 'text-slate-600 hover:bg-slate-100 hover:text-slate-900'
                      )}
                      title={isCollapsed ? item.label : undefined}
                    >
                      <item.icon
                        className={cn(
                          'w-5 h-5 shrink-0',
                          active ? 'text-emerald-600' : 'text-slate-400'
                        )}
                      />
                      {!isCollapsed && (
                        <>
                          <span className="flex-1 text-left">{item.label}</span>
                          {item.badge && (
                            <Badge
                              variant={item.badgeVariant || 'secondary'}
                              className={cn(
                                'h-5 min-w-5 px-1.5 text-xs font-semibold',
                                item.badgeVariant === 'destructive'
                                  ? 'bg-red-500 text-white'
                                  : 'bg-emerald-100 text-emerald-700'
                              )}
                            >
                              {item.badge}
                            </Badge>
                          )}
                        </>
                      )}
                      {isCollapsed && item.badge && (
                        <span className={cn(
                          'absolute top-1 right-1 w-2 h-2 rounded-full',
                          item.badgeVariant === 'destructive' ? 'bg-red-500' : 'bg-emerald-500'
                        )} />
                      )}
                    </button>
                  )
                })}
              </div>
            </div>
          ))}
        </nav>
      </ScrollArea>

      {/* Footer */}
      {!isCollapsed && (
        <div className="px-4 py-3 border-t border-slate-200/80">
          <div className="flex items-center gap-2 px-2 py-1.5 text-xs text-slate-400">
            <GraduationCap className="w-3.5 h-3.5" />
            <span>StudentHire v1.0</span>
          </div>
        </div>
      )}
    </div>
  )

  // Desktop Sidebar
  return (
    <>
      {/* Desktop Sidebar */}
      <aside
        className={cn(
          'hidden lg:flex flex-col border-r border-slate-200 bg-white transition-all duration-300 shrink-0',
          isCollapsed ? 'w-[68px]' : 'w-64'
        )}
      >
        {sidebarContent}

        {/* Collapse Toggle (Desktop) */}
        <button
          onClick={onToggleCollapse}
          className="hidden lg:flex items-center justify-center h-10 border-t border-slate-200/80 text-slate-400 hover:text-slate-600 hover:bg-slate-50 transition-colors"
          aria-label={isCollapsed ? 'Expand sidebar' : 'Collapse sidebar'}
        >
          <ChevronLeft
            className={cn(
              'w-4 h-4 transition-transform duration-300',
              isCollapsed && 'rotate-180'
            )}
          />
        </button>
      </aside>

      {/* Mobile Sidebar Drawer */}
      <AnimatePresence>
        {isOpen && (
          <>
            {/* Backdrop */}
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-40 bg-black/50 lg:hidden"
              onClick={onClose}
            />

            {/* Drawer */}
            <motion.aside
              initial={{ x: -280 }}
              animate={{ x: 0 }}
              exit={{ x: -280 }}
              transition={{ type: 'spring', damping: 25, stiffness: 300 }}
              className="fixed left-0 top-0 bottom-0 z-50 w-72 bg-white shadow-2xl lg:hidden"
            >
              <div className="absolute right-3 top-4 z-10">
                <button
                  onClick={onClose}
                  className="w-8 h-8 rounded-lg flex items-center justify-center text-slate-400 hover:text-slate-600 hover:bg-slate-100 transition-colors"
                  aria-label="Close sidebar"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              {sidebarContent}
            </motion.aside>
          </>
        )}
      </AnimatePresence>
    </>
  )
}
