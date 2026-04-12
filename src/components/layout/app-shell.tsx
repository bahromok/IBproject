'use client'

import { useState } from 'react'
import { AnimatePresence, motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Sidebar } from './sidebar'
import { TopNav } from './top-nav'
import {
  GraduationCap,
} from 'lucide-react'

// ─── Direct imports (SPA is already client-side, no benefit to lazy loading) ───
import { LandingPage } from '@/components/landing/landing-page'
import { LoginForm } from '@/components/auth/login-form'
import { RegisterForm } from '@/components/auth/register-form'
import JobsMarketplace from '@/components/pages/jobs-marketplace'
import JobDetail from '@/components/pages/job-detail'
import PostJobForm from '@/components/pages/post-job-form'
import { MyJobs } from '@/components/pages/my-jobs'
import BrowseFreelancers from '@/components/pages/browse-freelancers'
import FreelancerProfile from '@/components/pages/freelancer-profile'
import { ClientDashboard } from '@/components/pages/client-dashboard'
import { FreelancerDashboard } from '@/components/pages/freelancer-dashboard'
import { MyProposals } from '@/components/pages/my-proposals'
import { MyContracts } from '@/components/pages/my-contracts'
import { ContractDetail } from '@/components/pages/contract-detail'
import { MessagesInbox } from '@/components/pages/messages-inbox'
import { MessagesConversation } from '@/components/pages/messages-conversation'
import { ProfileEditor } from '@/components/pages/profile-editor'
import { ReviewsPage } from '@/components/pages/reviews-page'
import { SavedJobs } from '@/components/pages/saved-jobs'
import { SettingsPage } from '@/components/pages/settings-page'
import { ReportForm } from '@/components/pages/report-form'
import { AdminDashboard } from '@/components/pages/admin-dashboard'
import { AdminUsers } from '@/components/pages/admin-users'
import { AdminJobs } from '@/components/pages/admin-jobs'
import { AdminReports } from '@/components/pages/admin-reports'
import { AdminDisputes } from '@/components/pages/admin-disputes'
import { AdminAnalytics } from '@/components/pages/admin-analytics'
import { AdminFraud } from '@/components/pages/admin-fraud'
import { NotificationsPage } from '@/components/pages/notifications-page'
import { PrivacyPage } from '@/components/pages/privacy-page'
import { TermsPage } from '@/components/pages/terms-page'
import { ForgotPasswordPage } from '@/components/pages/forgot-password-page'
import { ResetPasswordPage } from '@/components/pages/reset-password-page'

/* ──────────── View Router ──────────── */

function ViewRouter() {
  const { currentPage } = useNavigationStore()
  const { user } = useAuthStore()

  // Dashboard route: redirect based on role
  if (currentPage === 'dashboard') {
    if (user?.role === 'ADMIN') return <AdminDashboard />
    if (user?.role === 'CLIENT') return <ClientDashboard />
    return <FreelancerDashboard />
  }

  switch (currentPage) {
    case 'landing': return <LandingPage />
    case 'auth/login': return <LoginForm />
    case 'auth/register': return <RegisterForm />
    case 'auth/forgot-password': return <ForgotPasswordPage />
    case 'auth/reset-password': return <ResetPasswordPage />
    case 'jobs': return <JobsMarketplace />
    case 'jobs/detail': return <JobDetail />
    case 'jobs/post': return <PostJobForm />
    case 'jobs/my-jobs': return <MyJobs />
    case 'freelancers': return <BrowseFreelancers />
    case 'freelancers/detail': return <FreelancerProfile />
    case 'proposals': return <MyProposals />
    case 'contracts': return <MyContracts />
    case 'contracts/detail': return <ContractDetail />
    case 'messages': return <MessagesInbox />
    case 'messages/conversation': return <MessagesConversation />
    case 'profile': return <ProfileEditor />
    case 'reviews': return <ReviewsPage />
    case 'saved-jobs': return <SavedJobs />
    case 'settings': return <SettingsPage />
    case 'reports': return <ReportForm />
    case 'notifications': return <NotificationsPage />
    case 'admin/dashboard': return <AdminDashboard />
    case 'admin/users': return <AdminUsers />
    case 'admin/jobs': return <AdminJobs />
    case 'admin/reports': return <AdminReports />
    case 'admin/disputes': return <AdminDisputes />
    case 'admin/analytics': return <AdminAnalytics />
    case 'admin/fraud': return <AdminFraud />
    case 'privacy': return <PrivacyPage />
    case 'terms': return <TermsPage />
    default: return <LandingPage />
  }
}

/* ──────────── App Shell ──────────── */

export function AppShell() {
  const { isAuthenticated } = useAuthStore()
  const { currentPage } = useNavigationStore()
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)

  const isAuthView = currentPage === 'auth/login' || currentPage === 'auth/register' || currentPage === 'auth/forgot-password' || currentPage === 'auth/reset-password'
  const isLandingView = currentPage === 'landing'
  const isFullPage = isAuthView || isLandingView

  // Dashboard layout for authenticated pages ONLY
  if (isAuthenticated && !isFullPage) {
    return (
      <div className="min-h-screen bg-slate-50 flex">
        <Sidebar
          isOpen={sidebarOpen}
          onClose={() => setSidebarOpen(false)}
          isCollapsed={sidebarCollapsed}
          onToggleCollapse={() => setSidebarCollapsed(!sidebarCollapsed)}
        />
        <div className="flex-1 flex flex-col min-w-0">
          <TopNav onToggleSidebar={() => setSidebarOpen(true)} />
          <main className="flex-1 p-4 lg:p-6 overflow-auto">
            <AnimatePresence mode="wait">
              <ViewRouter />
            </AnimatePresence>
          </main>
        </div>
      </div>
    )
  }

  // Non-authenticated users see full-page views for everything
  return (
    <AnimatePresence mode="wait">
      <ViewRouter />
    </AnimatePresence>
  )
}
