'use client'

import { useState, useEffect } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency, timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import {
  FileText, Search, XCircle, Eye,
  Clock, Briefcase
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'

interface Proposal {
  id: string
  proposedAmount: number
  proposedDuration: string | null
  coverLetter: string
  status: string
  createdAt: string
  updatedAt: string
  job: {
    id: string
    title: string
    status: string
    budgetMin: number | null
    budgetMax: number | null
    budgetType: string
    client: { id: string; displayName: string; avatar: string | null; rating: number }
  }
}

interface ProposalSummary {
  total: number
  pending: number
  accepted: number
  rejected: number
}

function StatusBadge({ status }: { status: string }) {
  const variants: Record<string, string> = {
    PENDING: 'bg-yellow-100 text-yellow-700 border-yellow-200',
    SHORTLISTED: 'bg-emerald-100 text-emerald-700 border-emerald-200',
    ACCEPTED: 'bg-green-100 text-green-700 border-green-200',
    REJECTED: 'bg-red-100 text-red-700 border-red-200',
    WITHDRAWN: 'bg-slate-100 text-slate-500 border-slate-200',
  }
  const label = status.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
  return (
    <Badge variant="outline" className={`text-xs font-medium ${variants[status] || 'bg-slate-100 text-slate-600'}`}>
      {label}
    </Badge>
  )
}

export function MyProposals() {
  const { user } = useAuthStore()
  const { navigate } = useNavigationStore()
  const { toast } = useToast()
  const [proposals, setProposals] = useState<Proposal[]>([])
  const [summary, setSummary] = useState<ProposalSummary>({ total: 0, pending: 0, accepted: 0, rejected: 0 })
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!user?.id) return
    authFetch('/api/proposals?limit=50')
      .then(r => r.json())
      .then(data => {
        setProposals(data?.proposals || [])
        setSummary(data?.summary || { total: 0, pending: 0, accepted: 0, rejected: 0 })
        setLoading(false)
      })
      .catch(() => {
        setProposals([])
        setSummary({ total: 0, pending: 0, accepted: 0, rejected: 0 })
        setLoading(false)
      })
  }, [user?.id])

  const handleWithdraw = async (proposalId: string) => {
    try {
      const res = await authFetch(`/api/proposals/${proposalId}`, { method: 'PATCH' })
      if (res.ok) {
        toast({ title: 'Proposal withdrawn', description: 'Your proposal has been withdrawn successfully.' })
        // Update local state
        setProposals(prev => prev.map(p => p.id === proposalId ? { ...p, status: 'WITHDRAWN' } : p))
        setSummary(prev => ({
          ...prev,
          pending: Math.max(0, prev.pending - 1),
          rejected: prev.rejected + 1,
        }))
      } else {
        const data = await res.json()
        toast({ title: 'Error', description: data.error || 'Failed to withdraw proposal', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to withdraw proposal', variant: 'destructive' })
    }
  }

  const filteredProposals = (tab: string) => {
    if (tab === 'all') return proposals
    if (tab === 'pending') return proposals.filter(p => p.status === 'PENDING' || p.status === 'SHORTLISTED')
    if (tab === 'accepted') return proposals.filter(p => p.status === 'ACCEPTED')
    if (tab === 'rejected') return proposals.filter(p => p.status === 'REJECTED' || p.status === 'WITHDRAWN')
    return proposals
  }

  const totalProposed = proposals.reduce((sum, p) => sum + p.proposedAmount, 0)

  const ProposalCard = ({ proposal }: { proposal: Proposal }) => (
    <Card className="border-slate-200 hover:shadow-sm transition-shadow">
      <CardContent className="p-4">
        <div className="flex items-start justify-between gap-4">
          <div className="flex-1 min-w-0">
            <div
              className="text-sm font-medium text-slate-900 hover:text-emerald-600 cursor-pointer truncate"
              onClick={() => navigate('jobs/detail', { id: proposal.job.id })}
            >
              {proposal.job.title}
            </div>
            <div className="flex items-center gap-2 mt-1">
              <span className="text-xs text-slate-500">{proposal.job.client?.displayName || 'Client'}</span>
              <span className="text-xs text-slate-300">&middot;</span>
              <span className="text-xs text-slate-400">{timeAgo(proposal.createdAt)}</span>
              <span className="text-xs text-slate-300">&middot;</span>
              <StatusBadge status={proposal.job.status} />
            </div>
            <div className="flex items-center gap-3 mt-3">
              <div className="text-sm font-semibold text-emerald-600">
                {formatCurrency(proposal.proposedAmount)}
              </div>
              {proposal.proposedDuration && (
                <span className="text-xs text-slate-400">{proposal.proposedDuration}</span>
              )}
              <StatusBadge status={proposal.status} />
            </div>
            {proposal.coverLetter && (
              <p className="text-xs text-slate-500 mt-2 line-clamp-2">{proposal.coverLetter}</p>
            )}
          </div>
          <div className="flex items-center gap-2 shrink-0">
            {proposal.status === 'PENDING' && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => handleWithdraw(proposal.id)}
                className="text-xs h-8 border-red-200 text-red-600 hover:bg-red-50"
              >
                <XCircle className="w-3 h-3 mr-1" />
                Withdraw
              </Button>
            )}
            <Button
              variant="outline"
              size="sm"
              onClick={() => navigate('jobs/detail', { id: proposal.job.id })}
              className="text-xs h-8 border-slate-200 text-slate-600 hover:bg-slate-50"
            >
              <Eye className="w-3 h-3 mr-1" />
              View Job
            </Button>
          </div>
        </div>
      </CardContent>
    </Card>
  )

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-40" />
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
        </div>
        <Skeleton className="h-10 w-64" />
        {[1, 2, 3].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">My Proposals</h1>
        <p className="text-slate-500 mt-1">Track all proposals you&apos;ve submitted.</p>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Total"
          value={summary.total}
          icon={FileText}
          iconColor="bg-slate-100 text-slate-600"
        />
        <StatCard
          label="Pending"
          value={summary.pending}
          icon={Clock}
          iconColor="bg-yellow-100 text-yellow-600"
        />
        <StatCard
          label="Accepted"
          value={summary.accepted}
          icon={Briefcase}
          iconColor="bg-emerald-100 text-emerald-600"
        />
        <StatCard
          label="Total Proposed"
          value={totalProposed}
          format="currency"
          icon={FileText}
          iconColor="bg-amber-100 text-amber-600"
        />
      </div>

      {/* Tabs */}
      <Tabs defaultValue="all">
        <TabsList className="bg-slate-100">
          <TabsTrigger value="all" className="text-sm">All ({proposals.length})</TabsTrigger>
          <TabsTrigger value="pending" className="text-sm">Pending ({summary.pending})</TabsTrigger>
          <TabsTrigger value="accepted" className="text-sm">Accepted ({summary.accepted})</TabsTrigger>
          <TabsTrigger value="rejected" className="text-sm">Rejected ({summary.rejected})</TabsTrigger>
        </TabsList>

        {['all', 'pending', 'accepted', 'rejected'].map(tab => (
          <TabsContent key={tab} value={tab}>
            {filteredProposals(tab).length === 0 ? (
              <Card className="border-slate-200">
                <CardContent className="p-12">
                  <EmptyState
                    icon={FileText}
                    title={tab === 'all' ? 'No proposals yet' : `No ${tab} proposals`}
                    description={tab === 'all'
                      ? "You haven't submitted any proposals yet. Start browsing jobs to find opportunities."
                      : `You don't have any ${tab} proposals at the moment.`}
                    actionLabel="Browse Jobs"
                    onAction={() => navigate('jobs')}
                  />
                </CardContent>
              </Card>
            ) : (
              <div className="space-y-3">
                {filteredProposals(tab).map(proposal => (
                  <ProposalCard key={proposal.id} proposal={proposal} />
                ))}
              </div>
            )}
          </TabsContent>
        ))}
      </Tabs>
    </motion.div>
  )
}
