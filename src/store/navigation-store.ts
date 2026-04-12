import { create } from 'zustand'

export type PageName =
  | 'landing'
  | 'auth/login'
  | 'auth/register'
  | 'auth/forgot-password'
  | 'auth/reset-password'
  | 'dashboard'
  | 'admin/dashboard'
  | 'admin/users'
  | 'admin/jobs'
  | 'admin/reports'
  | 'admin/disputes'
  | 'admin/fraud'
  | 'admin/analytics'
  | 'jobs'
  | 'jobs/detail'
  | 'jobs/post'
  | 'jobs/my-jobs'
  | 'freelancers'
  | 'freelancers/detail'
  | 'contracts'
  | 'contracts/detail'
  | 'proposals'
  | 'messages'
  | 'messages/conversation'
  | 'profile'
  | 'reviews'
  | 'saved-jobs'
  | 'settings'
  | 'notifications'
  | 'reports'
  | 'privacy'
  | 'terms'

interface NavigationState {
  currentPage: PageName
  pageParams: Record<string, string>
  history: Array<{ page: PageName; params: Record<string, string> }>
  navigate: (page: PageName, params?: Record<string, string>) => void
  goBack: () => void
}

export const useNavigationStore = create<NavigationState>((set, get) => ({
  currentPage: 'landing',
  pageParams: {},
  history: [],

  navigate: (page: PageName, params?: Record<string, string>) => {
    const { currentPage, pageParams, history } = get()
    if (currentPage !== page || JSON.stringify(pageParams) !== JSON.stringify(params || {})) {
      set({
        currentPage: page,
        pageParams: params || {},
        history: [...history.slice(-49), { page: currentPage, params: pageParams }],
      })
    }
    if (typeof window !== 'undefined') {
      window.scrollTo({ top: 0, behavior: 'smooth' })
    }
  },

  goBack: () => {
    const { history } = get()
    if (history.length > 0) {
      const prev = history[history.length - 1]
      set({
        currentPage: prev.page,
        pageParams: prev.params,
        history: history.slice(0, -1),
      })
      if (typeof window !== 'undefined') {
        window.scrollTo({ top: 0, behavior: 'smooth' })
      }
    }
  },
}))
