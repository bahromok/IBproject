import { useState } from 'react'
import { motion } from 'framer-motion'
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { ArrowLeft, Mail, CheckCircle2 } from 'lucide-react'

export function ForgotPasswordPage() {
    const { navigate } = useNavigationStore()
    const [email, setEmail] = useState('')
    const [submitted, setSubmitted] = useState(false)

    const handleSubmit = (e: React.FormEvent) => {
        e.preventDefault()
        if (!email) return
        setSubmitted(true)
    }

    return (
        <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
            <div className="w-full max-w-md">
                <Button variant="ghost" onClick={() => navigate('auth/login')} className="mb-6 pl-0 text-slate-500 hover:bg-slate-100 flex items-center gap-2">
                    <ArrowLeft className="w-4 h-4" />
                    Back to Login
                </Button>

                <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }}>
                    <Card className="border-0 shadow-xl overflow-hidden">
                        {!submitted ? (
                            <form onSubmit={handleSubmit}>
                                <CardHeader className="pb-4">
                                    <div className="w-12 h-12 bg-emerald-100 rounded-xl flex items-center justify-center mb-4">
                                        <Mail className="w-6 h-6 text-emerald-600" />
                                    </div>
                                    <CardTitle className="text-2xl font-bold text-slate-900">Forgot Password</CardTitle>
                                    <CardDescription className="text-slate-500 mt-2 text-sm leading-relaxed">
                                        Enter the email address associated with your account, and we'll send you a link to reset your password.
                                    </CardDescription>
                                </CardHeader>
                                <CardContent className="space-y-6">
                                    <div className="space-y-2">
                                        <Label htmlFor="email" className="text-sm font-medium text-slate-700">Email Address</Label>
                                        <Input
                                            id="email"
                                            type="email"
                                            placeholder="name@example.com"
                                            value={email}
                                            onChange={(e) => setEmail(e.target.value)}
                                            required
                                            className="h-11"
                                        />
                                    </div>
                                    <Button type="submit" className="w-full h-11 bg-emerald-500 hover:bg-emerald-600 text-white font-medium shadow-sm transition-all shadow-emerald-500/20 hover:shadow-emerald-500/30">
                                        Send Reset Link
                                    </Button>
                                </CardContent>
                            </form>
                        ) : (
                            <div className="p-8 flex flex-col items-center text-center">
                                <div className="w-16 h-16 bg-emerald-100 rounded-full flex items-center justify-center mb-6">
                                    <CheckCircle2 className="w-8 h-8 text-emerald-600" />
                                </div>
                                <h2 className="text-2xl font-bold text-slate-900 mb-2">Check your inbox</h2>
                                <p className="text-slate-500 mb-8 max-w-sm mx-auto">
                                    We've sent a password reset link to <span className="font-medium text-slate-900">{email}</span>. Click the link to create a new password.
                                </p>
                                <Button variant="outline" onClick={() => navigate('auth/login')} className="w-full h-11 font-medium border-slate-200">
                                    Return to Login
                                </Button>
                            </div>
                        )}
                    </Card>
                </motion.div>
            </div>
        </div>
    )
}
