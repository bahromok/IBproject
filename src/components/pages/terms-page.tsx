import { motion } from 'framer-motion'
import { Card, CardContent } from '@/components/ui/card'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { ArrowLeft } from 'lucide-react'

export function TermsPage() {
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
                            <h1 className="text-3xl font-bold text-slate-900 mb-6">Terms of Service</h1>
                            <p className="text-slate-600 mb-4">Last Updated: April 2026</p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">1. Acceptance of Terms</h2>
                            <p className="text-slate-600 leading-relaxed">
                                By accessing StudentHire, you agree to be bound by these local laws regarding digital contracting and copyright retention. Any breach of terms will result in account suspension.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">2. Payments and Escrow</h2>
                            <p className="text-slate-600 leading-relaxed">
                                All fixed-price transactions are managed via secure milestone thresholds. Funds must be deposited prior to work beginning and will only be released upon mutual agreement.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">3. Prohibited Conduct</h2>
                            <p className="text-slate-600 leading-relaxed">
                                Users may not circumvent our platform to process payments externally, submit fraudulent identities, nor harass other members. We use AI monitoring to detect and pause questionable activity.
                            </p>

                            <h2 className="text-xl font-semibold text-slate-800 mt-8 mb-4">4. Dispute Resolution</h2>
                            <p className="text-slate-600 leading-relaxed">
                                If a dispute arises over completed milestones, the funds will remain securely in escrow while our support team reviews the communication history and milestone deliverables.
                            </p>
                        </CardContent>
                    </Card>
                </motion.div>
            </div>
        </div>
    )
}
