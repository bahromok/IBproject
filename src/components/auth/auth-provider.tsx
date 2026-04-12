'use client'

import { useEffect, type ReactNode } from 'react'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'

interface AuthProviderProps {
  children: ReactNode
}

export function AuthProvider({ children }: AuthProviderProps) {
  const { isAuthenticated, isLoading } = useAuthStore()
  const { currentPage, navigate } = useNavigationStore()

  // Protected views that require authentication
  const protectedViews = [
    'dashboard', 'jobs/post', 'jobs/my-jobs', 'proposals',
    'contracts', 'contracts/detail', 'messages', 'messages/conversation',
    'profile', 'reviews', 'settings', 'reports', 'saved-jobs',
    'admin/dashboard', 'admin/users', 'admin/jobs', 'admin/reports',
    'admin/disputes', 'admin/analytics',
  ]

  const isProtected = protectedViews.some(v => currentPage.startsWith(v))
  const { fetchSession } = useAuthStore()

  useEffect(() => {
    fetchSession()
  }, [fetchSession])

  useEffect(() => {
    if (!isLoading) {
      if (isProtected && !isAuthenticated) {
        navigate('auth/login')
      }

      // Redirect authenticated users away from auth pages
      if (isAuthenticated && (currentPage === 'auth/login' || currentPage === 'auth/register')) {
        navigate('dashboard')
      }

      const { user } = useAuthStore.getState()
      const adminViews = ['admin/dashboard', 'admin/users', 'admin/jobs', 'admin/reports', 'admin/disputes', 'admin/analytics']
      const isAdminView = adminViews.some(v => currentPage.startsWith(v))
      if (isAdminView && isAuthenticated && user?.role !== 'ADMIN') {
        navigate('dashboard')
      }
    }
  }, [currentPage, isLoading, isAuthenticated, navigate, isProtected])

  return <>{children}</>
}
