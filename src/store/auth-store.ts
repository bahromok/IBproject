import { create } from 'zustand'
import { authClient } from '@/lib/auth/client'
import { switchUserRole } from '@/app/actions/user'

export type UserRole = 'CLIENT' | 'FREELANCER' | 'ADMIN'

export interface User {
  id: string
  email: string
  role?: string
  name?: string | null
  displayName?: string
  emailVerified?: boolean
}

interface AuthState {
  user: User | null
  isAuthenticated: boolean
  isLoading: boolean
  setUser: (user: User | null) => void
  fetchSession: () => Promise<void>
  logout: () => Promise<void>
  login: (email: string, password: string) => Promise<{ success: boolean; error?: string }>
  register: (data: any) => Promise<{ success: boolean; error?: string; user?: any }>
  forgotPassword: (email: string) => Promise<{ success: boolean; error?: string }>
  resetPassword: (password: string) => Promise<{ success: boolean; error?: string }>
  switchRole: () => Promise<void>
}

export const useAuthStore = create<AuthState>((set, get) => ({
  user: null,
  isAuthenticated: false,
  isLoading: false,

  setUser: (user) => {
    set({ user, isAuthenticated: !!user, isLoading: false })
  },

  fetchSession: async () => {
    try {
      set({ isLoading: true })
      const { data } = await authClient.getSession()
      if (data?.user) {
        set({ user: data.user as User, isAuthenticated: true })
      } else {
        set({ user: null, isAuthenticated: false })
      }
    } catch {
      set({ user: null, isAuthenticated: false })
    } finally {
      set({ isLoading: false })
    }
  },

  login: async (email, password) => {
    set({ isLoading: true })
    try {
      const { data, error } = await (authClient as any).signIn.email({ email, password })
      if (error) {
        set({ isLoading: false })
        return { success: false, error: error.message || 'Invalid credentials' }
      }

      await get().fetchSession()
      return { success: true }
    } catch {
      set({ isLoading: false })
      return { success: false, error: 'Network error. Please try again.' }
    }
  },

  register: async (data) => {
    set({ isLoading: true })
    try {
      const { data: authData, error } = await (authClient as any).signUp.email({
        email: data.email,
        password: data.password,
        name: data.displayName || data.email.split('@')[0],
        role: data.role || 'FREELANCER',
      } as any) // Cast as any because better-auth types might not see the injected field yet

      if (error) {
        set({ isLoading: false })
        return { success: false, error: error.message || 'Registration failed' }
      }

      await get().fetchSession()
      set({ isLoading: false })
      return { success: true, user: authData?.user }
    } catch {
      set({ isLoading: false })
      return { success: false, error: 'Network error' }
    }
  },

  logout: async () => {
    await (authClient as any).signOut()
    set({ user: null, isAuthenticated: false, isLoading: false })
  },

  forgotPassword: async (email) => {
    set({ isLoading: true })
    try {
      const { error } = await (authClient as any).forgotPassword({
        email,
        redirectTo: '/auth/reset-password',
      })
      if (error) {
        set({ isLoading: false })
        return { success: false, error: error.message || 'Failed to send reset email' }
      }
      set({ isLoading: false })
      return { success: true }
    } catch {
      set({ isLoading: false })
      return { success: false, error: 'Network error' }
    }
  },

  resetPassword: async (newPassword) => {
    set({ isLoading: true })
    try {
      const { error } = await (authClient as any).resetPassword({
        newPassword,
      })
      if (error) {
        set({ isLoading: false })
        return { success: false, error: error.message || 'Failed to reset password' }
      }
      set({ isLoading: false })
      return { success: true }
    } catch {
      set({ isLoading: false })
      return { success: false, error: 'Network error' }
    }
  },

  switchRole: async () => {
    const { user } = get()
    if (!user) return

    const newRole = user.role === 'CLIENT' ? 'FREELANCER' : 'CLIENT'
    set({ isLoading: true })
    try {
      await switchUserRole(newRole as 'CLIENT' | 'FREELANCER')
      await get().fetchSession()
    } catch (error) {
      console.error('Failed to switch role:', error)
    } finally {
      set({ isLoading: false })
    }
  },
}))
