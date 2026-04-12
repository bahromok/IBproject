'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Progress } from '@/components/ui/progress'
import { Separator } from '@/components/ui/separator'
import { Skeleton } from '@/components/ui/skeleton'
import { Textarea } from '@/components/ui/textarea'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter, DialogDescription
} from '@/components/ui/dialog'
import {
  AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent,
  AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger
} from '@/components/ui/alert-dialog'
import {
  ArrowLeft, Calendar, CheckCircle2, Clock, DollarSign, Eye,
  Flag, FileText, MessageSquare, PlusCircle, Shield,
  Star, XCircle, ChevronRight, Play, Send, ThumbsUp, ThumbsDown, Wallet
} from 'lucide-react'
import { useSendTransaction, useWaitForTransactionReceipt } from 'wagmi'
import { parseEther } from 'viem'
import { ConnectKitButton } from 'connectkit'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'
import { recordCryptoTransaction } from '@/app/actions/payments'
import { useSearchParams } from 'next/navigation'
import { useAccount } from 'wagmi'

interface Milestone {
  id: string
  title: string
  description: string | null
  amount: number
  status: string
  dueDate: string | null
  submissionNotes: string | null
  approvalNotes: string | null
  submittedAt: string | null
  approvedAt: string | null
  orderIndex: number
  createdAt: string
  transactions: Array<{ id: string; type: string; amount: number; status: string }>
}

interface Transaction {
  id: string
  type: string
  amount: number
  status: string
  description: string | null
  createdAt: string
}

interface ContractData {
  id: string
  title: string
  description: string | null
  contractType: string
  totalAmount: number
  hourlyRate: number | null
  status: string
  startDate: string | null
  endDate: string | null
  createdAt: string
  client: { id: string; displayName: string; avatar: string | null; location: string | null }
  freelancer: { id: string; displayName: string; avatar: string | null; title: string | null; location: string | null }
  job: { id: string; title: string; category: string }
  milestones: Milestone[]
  transactions: Transaction[]
  reviews: any[]
}

function MilestoneStatusBadge({ status }: { status: string }) {
  const config: Record<string, { color: string; icon: React.ElementType }> = {
    PENDING: { color: 'bg-slate-100 text-slate-600 border-slate-200', icon: Clock },
    IN_PROGRESS: { color: 'bg-yellow-100 text-yellow-700 border-yellow-200', icon: Play },
    SUBMITTED: { color: 'bg-orange-100 text-orange-700 border-orange-200', icon: Send },
    REQUESTED_APPROVAL: { color: 'bg-teal-100 text-teal-700 border-teal-200', icon: Eye },
    APPROVED: { color: 'bg-emerald-100 text-emerald-700 border-emerald-200', icon: ThumbsUp },
    PAID: { color: 'bg-green-100 text-green-700 border-green-200', icon: CheckCircle2 },
    REJECTED: { color: 'bg-red-100 text-red-700 border-red-200', icon: ThumbsDown },
  }
  const c = config[status] || config.PENDING
  const Icon = c.icon
  return (
    <Badge variant="outline" className={`text-xs font-medium gap-1 ${c.color}`}>
      <Icon className="w-3 h-3" />
      {status.replace(/_/g, ' ')}
    </Badge>
  )
}

function ContractStatusBadge({ status }: { status: string }) {
  const variants: Record<string, string> = {
    ACTIVE: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    PAUSED: 'bg-amber-100 text-amber-700 border-amber-200',
    COMPLETED: 'bg-green-100 text-green-700 border-green-200',
    CANCELLED: 'bg-slate-100 text-slate-500 border-slate-200',
    DISPUTED: 'bg-red-100 text-red-700 border-red-200',
  }
  return (
    <Badge variant="outline" className={`text-xs font-medium ${variants[status] || 'bg-slate-100 text-slate-600'}`}>
      {status}
    </Badge>
  )
}

export function ContractDetail() {
  const { user } = useAuthStore()
  const { pageParams, navigate, goBack } = useNavigationStore()
  const { toast } = useToast()
  const [contract, setContract] = useState<ContractData | null>(null)
  const [loading, setLoading] = useState(true)
  const [submitDialogOpen, setSubmitDialogOpen] = useState(false)
  const [submitNotes, setSubmitNotes] = useState('')
  const [submitMilestoneId, setSubmitMilestoneId] = useState('')
  const [addMilestoneOpen, setAddMilestoneOpen] = useState(false)
  const [newMilestone, setNewMilestone] = useState({ title: '', description: '', amount: '', dueDate: '' })
  const [feedbackDialogOpen, setFeedbackDialogOpen] = useState(false)
  const [feedbackAction, setFeedbackAction] = useState<'approve' | 'reject'>('approve')
  const [feedbackMilestoneId, setFeedbackMilestoneId] = useState('')
  const [feedbackNotes, setFeedbackNotes] = useState('')
  const [fundingMilestoneId, setFundingMilestoneId] = useState<string | null>(null)
  const { address } = useAccount()

  const searchParams = useSearchParams()

  const contractId = pageParams?.id

  const fetchContract = useCallback(async () => {
    if (!contractId) return
    setLoading(true)
    try {
      const res = await authFetch(`/api/contracts/${contractId}`)
      if (res.ok) {
        const data = await res.json()
        setContract(data.contract)
      } else {
        toast({ title: 'Error', description: 'Contract not found', variant: 'destructive' })
        navigate('contracts')
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to load contract', variant: 'destructive' })
    } finally {
      setLoading(false)
    }
  }, [contractId, toast, navigate])

  useEffect(() => { fetchContract() }, [fetchContract])

  const isClient = user?.id === contract?.client?.id
  const isFreelancer = user?.id === contract?.freelancer?.id
  const otherParty = isClient ? contract?.freelancer : contract?.client

  // Milestone actions
  const handleStartMilestone = async (milestoneId: string) => {
    try {
      const res = await authFetch(`/api/contracts/${contractId}/milestones/${milestoneId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'in_progress' }),
      })
      if (res.ok) {
        toast({ title: 'Milestone started', description: 'You can now start working on this milestone.' })
        fetchContract()
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to start milestone', variant: 'destructive' })
    }
  }

  const handleSubmitMilestone = async () => {
    try {
      const res = await authFetch(`/api/contracts/${contractId}/milestones/${submitMilestoneId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'submitted', submissionNote: submitNotes }),
      })
      if (res.ok) {
        toast({ title: 'Milestone submitted', description: 'Your work has been submitted for review.' })
        setSubmitDialogOpen(false)
        setSubmitNotes('')
        fetchContract()
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to submit milestone', variant: 'destructive' })
    }
  }

  const handleApproveReject = async () => {
    try {
      const res = await authFetch(`/api/contracts/${contractId}/milestones/${feedbackMilestoneId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: feedbackAction, feedback: feedbackNotes }),
      })
      if (res.ok) {
        toast({
          title: feedbackAction === 'approve' ? 'Milestone approved' : 'Milestone rejected',
          description: feedbackAction === 'approve' ? 'Payment will be released to the freelancer.' : 'The freelancer has been notified to make changes.',
        })
        setFeedbackDialogOpen(false)
        setFeedbackNotes('')
        fetchContract()
      } else {
        const data = await res.json()
        toast({ title: 'Error', description: data.error || 'Failed to process', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to process milestone', variant: 'destructive' })
    }
  }

  const handleAddMilestone = async () => {
    try {
      const res = await authFetch(`/api/contracts/${contractId}/milestones`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newMilestone),
      })
      if (res.ok) {
        toast({ title: 'Milestone added', description: 'A new milestone has been created.' })
        setAddMilestoneOpen(false)
        setNewMilestone({ title: '', description: '', amount: '', dueDate: '' })
        fetchContract()
      } else {
        const data = await res.json()
        toast({ title: 'Error', description: data.error || 'Failed to add milestone', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to add milestone', variant: 'destructive' })
    }
  }

  const handleUpdateContractStatus = async (status: string) => {
    try {
      const res = await authFetch(`/api/contracts/${contractId}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
      })
      if (res.ok) {
        toast({ title: 'Contract updated', description: `Contract status changed to ${status}.` })
        fetchContract()
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to update contract', variant: 'destructive' })
    }
  }

  const { sendTransaction, data: hash, isPending: isTxPending } = useSendTransaction()
  const { isLoading: isConfirming, isSuccess: isConfirmed } = useWaitForTransactionReceipt({ hash })

  const handleFundMilestone = async (milestoneId: string, method: 'stripe' | 'crypto', amount: number) => {
    if (method === 'stripe') {
      try {
        const res = await authFetch('/api/payments/stripe/checkout', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ milestoneId }),
        })
        if (res.ok) {
          const { url } = await res.json()
          window.location.href = url
        } else {
          const error = await res.json()
          toast({ title: 'Stripe Error', description: error.message || 'Failed to start checkout', variant: 'destructive' })
        }
      } catch (err) {
        toast({ title: 'Stripe Error', description: 'Network error occurred', variant: 'destructive' })
      }
    } else {
      // Real Crypto Transaction
      try {
        setFundingMilestoneId(milestoneId)
        // Platform or Freelancer address - in a real app this would come from the database
        const recipientAddress = '0x0000000000000000000000000000000000000000' // Placeholder for platform wallet

        sendTransaction({
          to: recipientAddress as `0x${string}`,
          value: parseEther((amount / 1000).toString()), // Demo conversion for ETH
        })

        toast({ title: 'Transaction Sent', description: 'Please wait for blockchain confirmation.' })
      } catch (err: any) {
        toast({ title: 'Wallet Error', description: err.message || 'Failed to send transaction', variant: 'destructive' })
      }
    }
  }

  // Handle Crypto Transaction Recording
  useEffect(() => {
    if (isConfirmed && hash && address && fundingMilestoneId) {
      const record = async () => {
        try {
          await recordCryptoTransaction(fundingMilestoneId, hash, address)
          toast({ title: 'Transaction Recorded', description: 'Your payment has been successfully logged.' })
          setFundingMilestoneId(null)
          fetchContract()
        } catch (err: any) {
          toast({ title: 'Recording Error', description: err.message, variant: 'destructive' })
        }
      }
      record()
    }
  }, [isConfirmed, hash, address, fundingMilestoneId, fetchContract, toast])

  // Computed values
  const milestones = contract?.milestones || []
  const totalMilestoneAmount = milestones.reduce((sum, m) => sum + m.amount, 0)
  const paidAmount = milestones.filter(m => m.status === 'PAID').reduce((sum, m) => sum + m.amount, 0)
  const approvedAmount = milestones.filter(m => m.status === 'APPROVED').reduce((sum, m) => sum + m.amount, 0)
  const releasedAmount = paidAmount + approvedAmount
  const remainingAmount = totalMilestoneAmount - releasedAmount

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-64" />
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 space-y-4">
            <Skeleton className="h-40 w-full rounded-xl" />
            <Skeleton className="h-60 w-full rounded-xl" />
          </div>
          <div className="space-y-4">
            <Skeleton className="h-40 w-full rounded-xl" />
            <Skeleton className="h-40 w-full rounded-xl" />
          </div>
        </div>
      </div>
    )
  }

  if (!contract) return null

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div className="flex items-start gap-4">
        <Button variant="ghost" size="icon" onClick={goBack} className="shrink-0 mt-1">
          <ArrowLeft className="w-5 h-5" />
        </Button>
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-3 flex-wrap">
            <h1 className="text-2xl font-bold text-slate-900">{contract.title}</h1>
            <ContractStatusBadge status={contract.status} />
          </div>
          <div className="flex items-center gap-4 mt-2 text-sm text-slate-500 flex-wrap">
            <span className="flex items-center gap-1">
              <Calendar className="w-4 h-4" />
              Started {contract.startDate ? new Date(contract.startDate).toLocaleDateString() : 'N/A'}
            </span>
            <span>•</span>
            <span>{contract.contractType === 'FIXED_PRICE' ? 'Fixed Price' : 'Hourly'}</span>
            <span>•</span>
            <span className="font-medium text-emerald-600">
              {contract.contractType === 'FIXED_PRICE'
                ? `$${contract.totalAmount.toLocaleString()}`
                : `$${contract.hourlyRate || 0}/hr`}
            </span>
          </div>
        </div>
        <div className="flex items-center gap-2 shrink-0">
          {contract.status === 'ACTIVE' && (
            <AlertDialog>
              <AlertDialogTrigger asChild>
                <Button variant="outline" size="sm" className="border-red-200 text-red-600 hover:bg-red-50">
                  <XCircle className="w-4 h-4 mr-1.5" />
                  End Contract
                </Button>
              </AlertDialogTrigger>
              <AlertDialogContent>
                <AlertDialogHeader>
                  <AlertDialogTitle>End this contract?</AlertDialogTitle>
                  <AlertDialogDescription>
                    This action will mark the contract as completed. Both parties will be able to leave reviews.
                  </AlertDialogDescription>
                </AlertDialogHeader>
                <AlertDialogFooter>
                  <AlertDialogCancel>Cancel</AlertDialogCancel>
                  <AlertDialogAction onClick={() => handleUpdateContractStatus('completed')} className="bg-emerald-500 hover:bg-emerald-600">
                    End Contract
                  </AlertDialogAction>
                </AlertDialogFooter>
              </AlertDialogContent>
            </AlertDialog>
          )}
          <Button variant="ghost" size="sm" onClick={() => navigate('messages')} className="text-slate-600">
            <MessageSquare className="w-4 h-4 mr-1.5" />
            Message
          </Button>
          <Button variant="ghost" size="sm" className="text-slate-600">
            <Flag className="w-4 h-4" />
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Main Content */}
        <div className="lg:col-span-2 space-y-6">
          {/* Other Party Info */}
          <Card className="border-slate-200">
            <CardContent className="p-4">
              <div className="flex items-center gap-3">
                <Avatar className="h-12 w-12 border border-slate-200">
                  <AvatarFallback className="bg-emerald-100 text-emerald-700 text-sm font-semibold">
                    {otherParty?.displayName?.[0] || '?'}
                  </AvatarFallback>
                </Avatar>
                <div className="flex-1">
                  <div className="text-sm font-semibold text-slate-900">{otherParty?.displayName || 'Anonymous'}</div>
                  {(otherParty as any)?.title && <div className="text-xs text-slate-500">{(otherParty as any).title}</div>}
                  {otherParty?.location && <div className="text-xs text-slate-400 mt-0.5">{otherParty.location}</div>}
                </div>
                <div className="text-xs text-slate-400 font-medium uppercase">
                  {isClient ? 'Freelancer' : 'Client'}
                </div>
              </div>
            </CardContent>
          </Card>

          {/* Milestones */}
          <Card className="border-slate-200">
            <CardHeader className="pb-3 flex flex-row items-center justify-between">
              <CardTitle className="text-lg font-semibold text-slate-900">
                Milestones ({milestones.length})
              </CardTitle>
              {isClient && contract.status === 'ACTIVE' && (
                <Button size="sm" onClick={() => setAddMilestoneOpen(true)} className="bg-emerald-500 hover:bg-emerald-600 text-white">
                  <PlusCircle className="w-4 h-4 mr-1.5" />
                  Add Milestone
                </Button>
              )}
            </CardHeader>
            <CardContent className="space-y-3">
              {milestones.length === 0 ? (
                <div className="py-8 text-center">
                  <FileText className="w-10 h-10 text-slate-300 mx-auto mb-3" />
                  <p className="text-sm text-slate-500">No milestones created yet</p>
                  {isClient && (
                    <Button size="sm" onClick={() => setAddMilestoneOpen(true)} variant="outline" className="mt-3 border-emerald-200 text-emerald-600">
                      <PlusCircle className="w-4 h-4 mr-1.5" />
                      Create First Milestone
                    </Button>
                  )}
                </div>
              ) : (
                milestones.map((milestone, idx) => (
                  <div key={milestone.id} className="border border-slate-200 rounded-xl p-4 hover:border-slate-300 transition-colors">
                    <div className="flex items-start justify-between gap-3">
                      <div className="flex items-start gap-3 flex-1 min-w-0">
                        <div className="w-8 h-8 rounded-lg bg-slate-100 flex items-center justify-center shrink-0 mt-0.5">
                          <span className="text-sm font-semibold text-slate-500">{idx + 1}</span>
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="text-sm font-semibold text-slate-900">{milestone.title}</div>
                          {milestone.description && (
                            <p className="text-xs text-slate-500 mt-1 line-clamp-2">{milestone.description}</p>
                          )}
                          <div className="flex items-center gap-3 mt-2 flex-wrap">
                            <MilestoneStatusBadge status={milestone.status} />
                            <span className="text-sm font-medium text-emerald-600">${milestone.amount.toLocaleString()}</span>
                            {milestone.dueDate && (
                              <span className="text-xs text-slate-400 flex items-center gap-1">
                                <Calendar className="w-3 h-3" />
                                Due {new Date(milestone.dueDate).toLocaleDateString()}
                              </span>
                            )}
                          </div>
                          {/* Submission notes */}
                          {milestone.submissionNotes && (
                            <div className="mt-3 p-3 bg-orange-50 rounded-lg border border-orange-100">
                              <div className="text-xs font-medium text-orange-700 mb-1">Submission Notes</div>
                              <p className="text-xs text-slate-600">{milestone.submissionNotes}</p>
                              {milestone.submittedAt && (
                                <div className="text-xs text-slate-400 mt-1">
                                  Submitted {new Date(milestone.submittedAt).toLocaleString()}
                                </div>
                              )}
                            </div>
                          )}
                          {/* Approval notes */}
                          {(milestone.approvalNotes || milestone.status === 'APPROVED') && (
                            <div className="mt-2 p-3 bg-emerald-50 rounded-lg border border-emerald-100">
                              <div className="text-xs font-medium text-emerald-700 mb-1">
                                {milestone.status === 'REJECTED' ? 'Rejection Reason' : 'Approval Notes'}
                              </div>
                              <p className="text-xs text-slate-600">{milestone.approvalNotes || 'Approved by client'}</p>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                    {/* Milestone action buttons */}
                    <div className="flex items-center gap-2 mt-3 ml-11">
                      {isClient && milestone.status === 'PENDING' && (
                        <div className="flex flex-wrap gap-2">
                          <Button
                            size="sm"
                            onClick={() => handleFundMilestone(milestone.id, 'stripe', milestone.amount)}
                            className="bg-indigo-500 hover:bg-indigo-600 text-white text-xs h-8"
                          >
                            <DollarSign className="w-3 h-3 mr-1" />
                            Fund with Stripe
                          </Button>
                          <div className="flex items-center gap-1">
                            <Button
                              size="sm"
                              onClick={() => handleFundMilestone(milestone.id, 'crypto', milestone.amount)}
                              className="bg-blue-500 hover:bg-blue-600 text-white text-xs h-8"
                              disabled={isTxPending || isConfirming}
                            >
                              <Wallet className="w-3 h-3 mr-1" />
                              {isTxPending || isConfirming ? 'Processing...' : 'Fund with Crypto'}
                            </Button>
                            <ConnectKitButton.Custom>
                              {({ isConnected, show }) => (
                                !isConnected && (
                                  <Button size="sm" variant="outline" onClick={show} className="text-[10px] h-8 border-dashed">
                                    Connect Wallet First
                                  </Button>
                                )
                              )}
                            </ConnectKitButton.Custom>
                          </div>
                        </div>
                      )}
                      {isFreelancer && milestone.status === 'PENDING' && (
                        <Button size="sm" onClick={() => handleStartMilestone(milestone.id)} className="bg-emerald-500 hover:bg-emerald-600 text-white text-xs h-8">
                          <Play className="w-3 h-3 mr-1" />
                          Start Working
                        </Button>
                      )}
                      {isFreelancer && milestone.status === 'IN_PROGRESS' && (
                        <Button size="sm" onClick={() => { setSubmitMilestoneId(milestone.id); setSubmitDialogOpen(true); }} className="bg-amber-500 hover:bg-amber-600 text-white text-xs h-8">
                          <Send className="w-3 h-3 mr-1" />
                          Submit for Review
                        </Button>
                      )}
                      {isClient && milestone.status === 'SUBMITTED' && (
                        <>
                          <Button size="sm" onClick={() => { setFeedbackMilestoneId(milestone.id); setFeedbackAction('approve'); setFeedbackDialogOpen(true); }} className="bg-emerald-500 hover:bg-emerald-600 text-white text-xs h-8">
                            <ThumbsUp className="w-3 h-3 mr-1" />
                            Approve
                          </Button>
                          <Button size="sm" onClick={() => { setFeedbackMilestoneId(milestone.id); setFeedbackAction('reject'); setFeedbackDialogOpen(true); }} variant="outline" className="border-red-200 text-red-600 hover:bg-red-50 text-xs h-8">
                            <ThumbsDown className="w-3 h-3 mr-1" />
                            Request Changes
                          </Button>
                        </>
                      )}
                      {milestone.status === 'PAID' && (
                        <Badge variant="outline" className="bg-green-50 text-green-600 border-green-200 gap-1">
                          <CheckCircle2 className="w-3 h-3" />
                          Paid
                        </Badge>
                      )}
                    </div>
                  </div>
                ))
              )}
            </CardContent>
          </Card>

          {/* Activity Log */}
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900">Activity Log</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-3">
                {/* Generate activity from milestones + transactions */}
                {milestones.map((m, idx) => (
                  <div key={`m-${m.id}`} className="flex items-start gap-3 pb-3">
                    <div className="w-2 h-2 rounded-full bg-emerald-400 mt-2 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <div className="text-sm text-slate-700">
                        <span className="font-medium">Milestone &quot;{m.title}&quot;</span> created
                      </div>
                      <div className="text-xs text-slate-400 mt-0.5">
                        ${m.amount.toLocaleString()} • {new Date(m.createdAt).toLocaleString()}
                      </div>
                    </div>
                  </div>
                ))}
                {contract.transactions.slice(0, 5).map((tx) => (
                  <div key={tx.id} className="flex items-start gap-3 pb-3">
                    <div className={`w-2 h-2 rounded-full mt-2 shrink-0 ${tx.type.includes('escrow') ? 'bg-amber-400' : 'bg-emerald-400'}`} />
                    <div className="flex-1 min-w-0">
                      <div className="text-sm text-slate-700">
                        <span className="font-medium">{tx.type.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase())}</span>
                        {tx.description && ` - ${tx.description}`}
                      </div>
                      <div className="text-xs text-slate-400 mt-0.5">
                        ${tx.amount.toLocaleString()} • {new Date(tx.createdAt).toLocaleString()}
                      </div>
                    </div>
                  </div>
                ))}
                {milestones.length === 0 && contract.transactions.length === 0 && (
                  <p className="text-sm text-slate-500 text-center py-4">No activity yet</p>
                )}
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Sidebar */}
        <div className="space-y-6">
          {/* Payment Summary */}
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <DollarSign className="w-5 h-5 text-emerald-500" />
                Payment Summary
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex justify-between items-center">
                <span className="text-sm text-slate-500">Total Contract Value</span>
                <span className="text-sm font-bold text-slate-900">${totalMilestoneAmount.toLocaleString()}</span>
              </div>
              <Separator />
              <div className="flex justify-between items-center">
                <span className="text-sm text-slate-500">Released</span>
                <span className="text-sm font-medium text-emerald-600">${releasedAmount.toLocaleString()}</span>
              </div>
              <div className="flex justify-between items-center">
                <span className="text-sm text-slate-500">Remaining</span>
                <span className="text-sm font-medium text-slate-900">${remainingAmount.toLocaleString()}</span>
              </div>
              <div className="pt-2">
                <div className="flex justify-between text-xs text-slate-500 mb-1.5">
                  <span>Progress</span>
                  <span>{totalMilestoneAmount > 0 ? Math.round((releasedAmount / totalMilestoneAmount) * 100) : 0}%</span>
                </div>
                <Progress value={totalMilestoneAmount > 0 ? (releasedAmount / totalMilestoneAmount) * 100 : 0} className="h-2" />
              </div>
            </CardContent>
          </Card>

          {/* Contract Details */}
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900">Contract Details</CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              {[
                { label: 'Type', value: contract.contractType === 'FIXED_PRICE' ? 'Fixed Price' : 'Hourly Rate' },
                { label: 'Budget', value: contract.contractType === 'FIXED_PRICE' ? `$${contract.totalAmount.toLocaleString()}` : `$${contract.hourlyRate || 0}/hr` },
                { label: 'Start Date', value: contract.startDate ? new Date(contract.startDate).toLocaleDateString() : 'N/A' },
                { label: 'End Date', value: contract.endDate ? new Date(contract.endDate).toLocaleDateString() : 'N/A' },
                { label: 'Milestones', value: `${milestones.filter(m => m.status === 'PAID').length}/${milestones.length} completed` },
              ].map(item => (
                <div key={item.label} className="flex justify-between">
                  <span className="text-xs text-slate-500">{item.label}</span>
                  <span className="text-xs font-medium text-slate-900">{item.value}</span>
                </div>
              ))}
            </CardContent>
          </Card>

          {/* Quick Actions */}
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900">Actions</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              <Button variant="outline" className="w-full justify-start text-sm border-slate-200 hover:bg-slate-50" onClick={() => navigate('messages')}>
                <MessageSquare className="w-4 h-4 mr-2 text-slate-400" />
                Send Message
              </Button>
              {contract.status === 'ACTIVE' && (
                <Button variant="outline" className="w-full justify-start text-sm border-amber-200 text-amber-600 hover:bg-amber-50" onClick={() => handleUpdateContractStatus('paused')}>
                  <Clock className="w-4 h-4 mr-2" />
                  Pause Contract
                </Button>
              )}
              {contract.status === 'PAUSED' && (
                <Button variant="outline" className="w-full justify-start text-sm border-emerald-200 text-emerald-600 hover:bg-emerald-50" onClick={() => handleUpdateContractStatus('ACTIVE')}>
                  <Play className="w-4 h-4 mr-2" />
                  Resume Contract
                </Button>
              )}
              <Button variant="outline" className="w-full justify-start text-sm border-red-200 text-red-600 hover:bg-red-50" onClick={() => handleUpdateContractStatus('disputed')}>
                <Shield className="w-4 h-4 mr-2" />
                File Dispute
              </Button>
            </CardContent>
          </Card>
        </div>
      </div>

      {/* Submit Milestone Dialog */}
      <Dialog open={submitDialogOpen} onOpenChange={setSubmitDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Submit Milestone for Review</DialogTitle>
            <DialogDescription>Add notes about your submission for the client to review.</DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-2">
            <div>
              <Label>Submission Notes</Label>
              <Textarea
                value={submitNotes}
                onChange={(e) => setSubmitNotes(e.target.value)}
                placeholder="Describe what you've completed..."
                className="mt-1.5 min-h-[100px]"
              />
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setSubmitDialogOpen(false)}>Cancel</Button>
            <Button onClick={handleSubmitMilestone} className="bg-emerald-500 hover:bg-emerald-600 text-white">
              <Send className="w-4 h-4 mr-1.5" />
              Submit
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Approve/Reject Dialog */}
      <Dialog open={feedbackDialogOpen} onOpenChange={setFeedbackDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {feedbackAction === 'approve' ? 'Approve Milestone' : 'Request Changes'}
            </DialogTitle>
            <DialogDescription>
              {feedbackAction === 'approve'
                ? 'Payment will be released to the freelancer.'
                : 'The freelancer will be asked to make revisions.'}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-2">
            <div>
              <Label>{feedbackAction === 'approve' ? 'Approval Notes (optional)' : 'Describe what needs to change'}</Label>
              <Textarea
                value={feedbackNotes}
                onChange={(e) => setFeedbackNotes(e.target.value)}
                placeholder={feedbackAction === 'approve' ? 'Add any notes...' : 'Describe the changes needed...'}
                className="mt-1.5 min-h-[100px]"
              />
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setFeedbackDialogOpen(false)}>Cancel</Button>
            <Button
              onClick={handleApproveReject}
              className={feedbackAction === 'approve' ? 'bg-emerald-500 hover:bg-emerald-600 text-white' : 'bg-red-500 hover:bg-red-600 text-white'}
            >
              {feedbackAction === 'approve' ? 'Approve & Release Payment' : 'Request Changes'}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Add Milestone Dialog */}
      <Dialog open={addMilestoneOpen} onOpenChange={setAddMilestoneOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Add New Milestone</DialogTitle>
            <DialogDescription>Create a new milestone for this contract.</DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-2">
            <div>
              <Label>Title *</Label>
              <Input
                value={newMilestone.title}
                onChange={(e) => setNewMilestone({ ...newMilestone, title: e.target.value })}
                placeholder="Milestone title"
                className="mt-1.5"
              />
            </div>
            <div>
              <Label>Description *</Label>
              <Textarea
                value={newMilestone.description}
                onChange={(e) => setNewMilestone({ ...newMilestone, description: e.target.value })}
                placeholder="Describe what needs to be delivered..."
                className="mt-1.5 min-h-[80px]"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label>Amount ($) *</Label>
                <Input
                  type="number"
                  value={newMilestone.amount}
                  onChange={(e) => setNewMilestone({ ...newMilestone, amount: e.target.value })}
                  placeholder="0.00"
                  className="mt-1.5"
                />
              </div>
              <div>
                <Label>Due Date</Label>
                <Input
                  type="date"
                  value={newMilestone.dueDate}
                  onChange={(e) => setNewMilestone({ ...newMilestone, dueDate: e.target.value })}
                  className="mt-1.5"
                />
              </div>
            </div>
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setAddMilestoneOpen(false)}>Cancel</Button>
            <Button onClick={handleAddMilestone} className="bg-emerald-500 hover:bg-emerald-600 text-white" disabled={!newMilestone.title || !newMilestone.description || !newMilestone.amount}>
              <PlusCircle className="w-4 h-4 mr-1.5" />
              Add Milestone
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </motion.div>
  )
}
