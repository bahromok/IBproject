'use client'

import { useState } from 'react'
import { motion } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { useToast } from '@/hooks/use-toast'
import { GraduationCap, Lock, Eye, EyeOff, Loader2, ArrowLeft } from 'lucide-react'

export function ResetPasswordForm() {
    const [password, setPassword] = useState('')
    const [confirmPassword, setConfirmPassword] = useState('')
    const [showPassword, setShowPassword] = useState(false)
    const [isSubmitting, setIsSubmitting] = useState(false)
    const { resetPassword } = useAuthStore()
    const { navigate } = useNavigationStore()
    const { toast } = useToast()

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault()

        if (!password) {
            toast({ title: 'Error', description: 'Please enter a new password.', variant: 'destructive' })
            return
        }
        if (password.length < 8) {
            toast({ title: 'Error', description: 'Password must be at least 8 characters.', variant: 'destructive' })
            return
        }
        if (password !== confirmPassword) {
            toast({ title: 'Error', description: 'Passwords do not match.', variant: 'destructive' })
            return
        }

        setIsSubmitting(true)
        const result = await resetPassword(password)

        if (result.success) {
            toast({ title: 'Password updated!', description: 'You can now sign in with your new password.' })
            navigate('auth/login')
        } else {
            toast({ title: 'Error', description: result.error || 'Failed to reset password', variant: 'destructive' })
        }
        setIsSubmitting(false)
    }

    return (
        <div className="min-h-screen flex items-center justify-center bg-slate-50 px-4">
            <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                className="w-full max-w-md"
            >
                <Card className="border-0 shadow-xl overflow-hidden">
                    <div className="h-2 bg-emerald-500" />
                    <CardHeader className="space-y-1">
                        <button
                            onClick={() => navigate('auth/login')}
                            className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700 transition-colors mb-4 w-fit"
                        >
                            <ArrowLeft className="w-3.5 h-3.5" />
                            Back to login
                        </button>
                        <div className="flex items-center gap-2 mb-2">
                            <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center">
                                <GraduationCap className="w-5 h-5 text-white" />
                            </div>
                            <span className="text-xl font-bold text-slate-900">StudentHire</span>
                        </div>
                        <CardTitle className="text-2xl font-bold">Set new password</CardTitle>
                        <CardDescription>
                            Your new password must be different from previous used passwords.
                        </CardDescription>
                    </CardHeader>

                    <CardContent>
                        <form onSubmit={handleSubmit} className="space-y-4">
                            <div className="space-y-2">
                                <Label htmlFor="password">New Password</Label>
                                <div className="relative">
                                    <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                                    <Input
                                        id="password"
                                        type={showPassword ? 'text' : 'password'}
                                        placeholder="Min 8 characters"
                                        value={password}
                                        onChange={(e) => setPassword(e.target.value)}
                                        className="pl-10 pr-10 h-11"
                                        disabled={isSubmitting}
                                    />
                                    <button
                                        type="button"
                                        onClick={() => setShowPassword(!showPassword)}
                                        className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 transition-colors"
                                    >
                                        {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                                    </button>
                                </div>
                            </div>

                            <div className="space-y-2">
                                <Label htmlFor="confirmPassword">Confirm Password</Label>
                                <div className="relative">
                                    <Lock className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                                    <Input
                                        id="confirmPassword"
                                        type={showPassword ? 'text' : 'password'}
                                        placeholder="Confirm new password"
                                        value={confirmPassword}
                                        onChange={(e) => setConfirmPassword(e.target.value)}
                                        className="pl-10 h-11"
                                        disabled={isSubmitting}
                                    />
                                </div>
                            </div>

                            <Button
                                type="submit"
                                className="w-full h-11 bg-emerald-500 hover:bg-emerald-600 text-white shadow-lg shadow-emerald-500/20"
                                disabled={isSubmitting}
                            >
                                {isSubmitting ? (
                                    <>
                                        <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                                        Updating password...
                                    </>
                                ) : (
                                    'Reset password'
                                )}
                            </Button>
                        </form>
                    </CardContent>
                </Card>
            </motion.div>
        </div>
    )
}
