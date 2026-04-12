'use client'
import { useState } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { useAuthStore } from '@/store/auth-store'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'
import { authClient } from '@/lib/auth/client'
import { Mail, CheckCircle, X, Loader2 } from 'lucide-react'

export function EmailVerificationBanner() {
  const { user } = useAuthStore()
  const { toast } = useToast()
  const [showForm, setShowForm] = useState(false)
  const [code, setCode] = useState('')
  const [sentCode, setSentCode] = useState('')
  const [loading, setLoading] = useState(false)
  const [verified, setVerified] = useState(false)
  const [dismissed, setDismissed] = useState(false)

  if (!user || user.emailVerified || dismissed) return null

  const handleSendCode = async () => {
    setLoading(true)
    try {
      const { error } = await (authClient as any).emailOtp.sendVerificationOtp({
        email: user.email,
        type: 'email-verification',
      })

      if (!error) {
        setShowForm(true)
        toast({ title: 'Verification code sent!', description: 'Please check your email inbox to proceed.' })
      } else {
        toast({ title: 'Error', description: error.message || 'Failed to dispatch email.', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to send verification', variant: 'destructive' })
    }
    setLoading(false)
  }

  const handleVerify = async () => {
    if (!code.trim()) {
      toast({ title: 'Error', description: 'Enter the verification code', variant: 'destructive' })
      return
    }
    setLoading(true)
    try {
      const { data, error } = await (authClient as any).emailOtp.verifyEmail({
        email: user.email,
        otp: code.trim(),
      })

      if (!error) {
        setVerified(true)
        // Refresh session using our store mapping
        await useAuthStore.getState().fetchSession()
        toast({ title: 'Email Verified! 🎉', description: 'Your email has been successfully registered on the Neon branch.' })
      } else {
        toast({ title: 'Verification failed', description: error.message || 'Invalid code.', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Verification failed', variant: 'destructive' })
    }
    setLoading(false)
  }

  return (
    <motion.div initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }} className="mb-4">
      <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
        <div className="flex items-start gap-3">
          <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center shrink-0 mt-0.5">
            {verified ? (
              <CheckCircle className="w-5 h-5 text-emerald-600" />
            ) : (
              <Mail className="w-5 h-5 text-amber-600" />
            )}
          </div>
          <div className="flex-1 min-w-0">
            {verified ? (
              <div>
                <p className="text-sm font-semibold text-emerald-800">Email Verified Successfully!</p>
                <p className="text-xs text-emerald-600 mt-0.5">
                  Your email {user.email} has been verified. You now have full access to all features.
                </p>
              </div>
            ) : (
              <div>
                <p className="text-sm font-semibold text-amber-800">Verify your email address</p>
                <p className="text-xs text-amber-600 mt-0.5">
                  Please verify <span className="font-medium">{user.email}</span> to unlock all features.
                  Unverified accounts have limited access.
                </p>
                <AnimatePresence>
                  {showForm ? (
                    <motion.div
                      initial={{ opacity: 0, height: 0 }}
                      animate={{ opacity: 1, height: 'auto' }}
                      exit={{ opacity: 0, height: 0 }}
                      className="mt-3"
                    >
                      <div className="flex gap-2">
                        <Input
                          placeholder="Enter verification code"
                          value={code}
                          onChange={(e) => setCode(e.target.value)}
                          className="h-9 text-sm max-w-[200px] font-mono"
                        />
                        <Button
                          size="sm"
                          onClick={handleVerify}
                          disabled={loading}
                          className="h-9 bg-emerald-500 hover:bg-emerald-600 text-white px-4"
                        >
                          {loading ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : 'Verify'}
                        </Button>
                      </div>
                      <p className="text-xs text-amber-500 mt-1.5">
                        Please check your email inbox or spam folder for your code.
                      </p>
                    </motion.div>
                  ) : (
                    <Button
                      size="sm"
                      onClick={handleSendCode}
                      disabled={loading}
                      className="mt-2 bg-amber-500 hover:bg-amber-600 text-white h-8 text-xs px-3"
                    >
                      {loading ? (
                        <Loader2 className="w-3 h-3 animate-spin" />
                      ) : (
                        <>
                          <Mail className="w-3 h-3 mr-1" /> Send Verification Code
                        </>
                      )}
                    </Button>
                  )}
                </AnimatePresence>
              </div>
            )}
          </div>
          <button
            onClick={() => setDismissed(true)}
            className="text-amber-400 hover:text-amber-600 transition-colors shrink-0"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      </div>
    </motion.div>
  )
}
