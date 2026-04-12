import { motion } from 'framer-motion'
import { Card, CardContent } from '@/components/ui/card'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { ArrowLeft } from 'lucide-react'

export function PrivacyPage() {
    const { navigate } = useNavigationStore()

    return (
        <div className="min-h-screen bg-slate-50 py-12 px-4 sm:px-6 lg:px-8">
            <div className="max-w-3xl mx-auto">
                <Button variant="ghost" onClick={() => navigate('landing')} className="mb-6 pl-0 hover:bg-slate-100 flex items-center gap-2 text-slate-500">
                    <ArrowLeft className="w-4 h-4" />
                    Back to Home
                </Button>
                <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }}>
                    <Card className="border-0 shadow-lg">
                        <CardContent className="p-8 md:p-12 prose prose-slate max-w-none">
                            <h1 className="text-3xl font-bold text-slate-900 mb-6">Privacy Policy</h1>
                            <p className="text-slate-600 mb-4">Last Updated: April 2026</p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">1. Information We Collect</h2>
                            <p className="text-slate-600 leading-relaxed">
                                When you register for an account, we collect basic information such as your email address, display name, and role constraints. For students, we may collect verification details relating to your institution.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">2. How We Use Your Information</h2>
                            <p className="text-slate-600 leading-relaxed">
                                We use the collected information only to power the StudentHire platform, process transaction flows safely via escrow, and generate analytics to optimize user experiences.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">3. Data Security</h2>
                            <p className="text-slate-600 leading-relaxed">
                                We implement a variety of strong security measures to protect against fraud, data leaks, and identity exposure. Real names and sensitive contacts are obscured from public view.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">4. Contact Us</h2>
                            <p className="text-slate-600 leading-relaxed">
                                If you have any questions or require data deletion, please contact our support team at support@studenthire.example.com.
                            </p>
                        </CardContent>
                    </Card>
                </motion.div>
            </div>
        </div>
    )
}
