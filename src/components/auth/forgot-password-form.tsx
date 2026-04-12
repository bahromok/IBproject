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
import { GraduationCap, Mail, Loader2, ArrowLeft, CheckCircle2 } from 'lucide-react'

export function ForgotPasswordForm() {
    const [email, setEmail] = useState('')
    const [isSubmitting, setIsSubmitting] = useState(false)
    const [isSent, setIsSent] = useState(false)
    const { forgotPassword } = useAuthStore()
    const { navigate } = useNavigationStore()
    const { toast } = useToast()

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault()

        if (!email.trim()) {
            toast({ title: 'Error', description: 'Please enter your email address.', variant: 'destructive' })
            return
        }

        setIsSubmitting(true)
        const result = await forgotPassword(email)

        if (result.success) {
            setIsSent(true)
            toast({ title: 'Reset link sent!', description: 'Please check your email for instructions.' })
        } else {
            toast({ title: 'Error', description: result.error || 'Failed to send reset link', variant: 'destructive' })
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
                        <CardTitle className="text-2xl font-bold">Forgot password?</CardTitle>
                        <CardDescription>
                            {isSent
                                ? "We've sent a recovery link to your email."
                                : "No worries, we'll send you reset instructions."}
                        </CardDescription>
                    </CardHeader>

                    <CardContent>
                        {isSent ? (
                            <div className="py-6 text-center space-y-4">
                                <div className="w-16 h-16 bg-emerald-50 rounded-full flex items-center justify-center mx-auto">
                                    <CheckCircle2 className="w-10 h-10 text-emerald-500" />
                                </div>
                                <p className="text-slate-600">
                                    Please check <span className="font-semibold text-slate-900">{email}</span> for a link to reset your password.
                                </p>
                                <Button
                                    variant="outline"
                                    className="w-full"
                                    onClick={() => setIsSent(false)}
                                >
                                    Didn't get the email? Try again
                                </Button>
                            </div>
                        ) : (
                            <form onSubmit={handleSubmit} className="space-y-4">
                                <div className="space-y-2">
                                    <Label htmlFor="email">Email address</Label>
                                    <div className="relative">
                                        <Mail className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                                        <Input
                                            id="email"
                                            type="email"
                                            placeholder="you@example.com"
                                            value={email}
                                            onChange={(e) => setEmail(e.target.value)}
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
                                            Sending link...
                                        </>
                                    ) : (
                                        'Send reset link'
                                    )}
                                </Button>
                            </form>
                        )}
                    </CardContent>
                </Card>
            </motion.div>
        </div>
    )
}
