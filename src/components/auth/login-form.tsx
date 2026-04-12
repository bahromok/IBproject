'use client'

import { useState } from 'react'
import { motion } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { useToast } from '@/hooks/use-toast'
import { GraduationCap, Eye, EyeOff, Loader2, ArrowLeft } from 'lucide-react'

export function LoginForm() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [isSubmitting, setIsSubmitting] = useState(false)
  const { login } = useAuthStore()
  const { navigate } = useNavigationStore()
  const { toast } = useToast()

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()

    if (!email.trim()) {
      toast({ title: 'Error', description: 'Please enter your email address.', variant: 'destructive' })
      return
    }
    if (!password) {
      toast({ title: 'Error', description: 'Please enter your password.', variant: 'destructive' })
      return
    }

    setIsSubmitting(true)
    const result = await login(email, password)

    if (result.success) {
      toast({ title: 'Welcome back!', description: 'You have signed in successfully.' })
      navigate('dashboard')
    } else {
      toast({ title: 'Sign in failed', description: result.error || 'Invalid credentials', variant: 'destructive' })
    }
    setIsSubmitting(false)
  }

  return (
    <div className="min-h-screen flex bg-white">
      {/* Left side - Decorative */}
      <div className="hidden lg:flex lg:w-1/2 bg-gradient-to-br from-emerald-500 via-emerald-600 to-teal-600 relative overflow-hidden">
        <div className="absolute inset-0">
          <div className="absolute top-20 left-20 w-72 h-72 bg-white/10 rounded-full blur-3xl" />
          <div className="absolute bottom-20 right-20 w-96 h-96 bg-teal-400/20 rounded-full blur-3xl" />
          <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-64 h-64 bg-emerald-300/10 rounded-full blur-2xl" />
        </div>

        <div className="relative z-10 flex flex-col justify-center px-16">
          <motion.div
            initial={{ opacity: 0, x: -30 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ duration: 0.7 }}
          >
            <div className="flex items-center gap-3 mb-8">
              <div className="w-12 h-12 rounded-xl bg-white/20 backdrop-blur-sm flex items-center justify-center">
                <GraduationCap className="w-7 h-7 text-white" />
              </div>
              <span className="text-2xl font-bold text-white">
                Student<span className="text-emerald-200">Hire</span>
              </span>
            </div>

            <h2 className="text-4xl font-bold text-white mb-4 leading-tight">
              Welcome back to<br />your workspace
            </h2>
            <p className="text-lg text-emerald-100 leading-relaxed max-w-md">
              Sign in to manage your projects, connect with clients or freelancers, and grow your career.
            </p>

            <div className="mt-12 space-y-4">
              {[
                { stat: '127', label: 'Active student freelancers' },
                { stat: '89', label: 'Projects completed this month' },
                { stat: '$48,200', label: 'Paid to freelancers this month' },
              ].map((item, i) => (
                <motion.div
                  key={item.label}
                  initial={{ opacity: 0, x: -20 }}
                  animate={{ opacity: 1, x: 0 }}
                  transition={{ delay: 0.3 + i * 0.1 }}
                  className="flex items-center gap-3"
                >
                  <div className="w-2 h-2 rounded-full bg-emerald-300" />
                  <span className="text-emerald-100">
                    <span className="font-semibold text-white">{item.stat}</span> {item.label}
                  </span>
                </motion.div>
              ))}
            </div>
          </motion.div>
        </div>
      </div>

      {/* Right side - Form */}
      <div className="flex-1 flex items-center justify-center px-4 sm:px-6 lg:px-8 py-12">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          className="w-full max-w-md"
        >
          <div className="lg:hidden flex items-center gap-2 mb-8 justify-center">
            <div className="w-9 h-9 rounded-lg bg-emerald-500 flex items-center justify-center">
              <GraduationCap className="w-5 h-5 text-white" />
            </div>
            <span className="text-xl font-bold text-slate-900">
              Student<span className="text-emerald-500">Hire</span>
            </span>
          </div>

          <Card className="border-0 shadow-xl">
            <CardHeader className="space-y-1 pb-2">
              <button
                onClick={() => navigate('landing')}
                className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700 transition-colors mb-2 w-fit"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                Back to home
              </button>
              <CardTitle className="text-2xl font-bold text-slate-900">Sign in</CardTitle>
              <CardDescription className="text-slate-500">
                Enter your credentials to access your account
              </CardDescription>
            </CardHeader>

            <CardContent className="pt-4">
              <form onSubmit={handleSubmit} className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="email" className="text-slate-700 font-medium">
                    Email address
                  </Label>
                  <Input
                    id="email"
                    type="email"
                    placeholder="you@example.com"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    autoComplete="email"
                    className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                    disabled={isSubmitting}
                  />
                </div>

                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <Label htmlFor="password" className="text-slate-700 font-medium">
                      Password
                    </Label>
                    <button
                      type="button"
                      onClick={() => navigate('auth/forgot-password')}
                      className="text-xs text-emerald-600 hover:text-emerald-700 font-medium"
                    >
                      Forgot password?
                    </button>
                  </div>
                  <div className="relative">
                    <Input
                      id="password"
                      type={showPassword ? 'text' : 'password'}
                      placeholder="Enter your password"
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      autoComplete="current-password"
                      className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20 pr-10"
                      disabled={isSubmitting}
                    />
                    <button
                      type="button"
                      onClick={() => setShowPassword(!showPassword)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors"
                      aria-label={showPassword ? 'Hide password' : 'Show password'}
                    >
                      {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                    </button>
                  </div>
                </div>

                <Button
                  type="submit"
                  className="w-full h-11 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg font-semibold text-base shadow-lg shadow-emerald-500/25 hover:shadow-emerald-500/40 transition-all"
                  disabled={isSubmitting}
                >
                  {isSubmitting ? (
                    <>
                      <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                      Signing in...
                    </>
                  ) : (
                    'Sign in'
                  )}
                </Button>
              </form>
            </CardContent>

            <CardFooter className="flex flex-col gap-4">
              <div className="text-sm text-center text-slate-500">
                Don&apos;t have an account?{' '}
                <button
                  onClick={() => navigate('auth/register')}
                  className="text-emerald-600 hover:text-emerald-700 font-semibold"
                >
                  Create one now
                </button>
              </div>
            </CardFooter>
          </Card>
        </motion.div>
      </div>
    </div>
  )
}
