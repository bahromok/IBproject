'use client'

import { useState, useEffect } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Progress } from '@/components/ui/progress'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Skeleton } from '@/components/ui/skeleton'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { formatCurrency, timeAgo } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import {
  Briefcase, ArrowRight, Clock, Eye, DollarSign,
} from 'lucide-react'

interface Contract {
  id: string
  title: string
  status: string
  contractType: string
  totalAmount: number
  hourlyRate: number | null
  createdAt: string
  updatedAt: string
  client: { id: string; displayName: string; avatar: string | null }
  freelancer: { id: string; displayName: string; avatar: string | null }
  job: { id: string; title: string }
  milestones: Array<{ id: string; title: string; amount: number; status: string }>
}

interface ApiTransaction {
  id: string
  type: string
  amount: number
  netAmount: number
  fromUserId: string
  toUserId: string
  status: string
  createdAt: string
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

export function MyContracts() {
  const { user } = useAuthStore()
  const { navigate } = useNavigationStore()
  const [contracts, setContracts] = useState<Contract[]>([])
  const [transactions, setTransactions] = useState<ApiTransaction[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!user?.id) return
    Promise.all([
      authFetch('/api/contracts?limit=50').then(r => r.json()).catch(() => ({})),
      authFetch('/api/transactions?limit=50').then(r => r.json()).catch(() => ({})),
    ]).then(([contractsData, txData]) => {
      setContracts(contractsData?.contracts || [])
      setTransactions(txData?.transactions || [])
      setLoading(false)
    })
  }, [user?.id])

  const activeCount = contracts.filter(c => c.status === 'ACTIVE' || c.status === 'PAUSED').length
  const completedCount = contracts.filter(c => c.status === 'COMPLETED').length

  // Compute spent/earned from transactions
  const totalSpent = transactions
    .filter(t => t.type === 'ESCROW_DEPOSIT' && t.fromUserId === user?.id && t.status === 'COMPLETED')
    .reduce((sum, t) => sum + t.amount, 0)
  const totalEarned = transactions
    .filter(t => t.type === 'PAYMENT_RELEASE' && t.toUserId === user?.id && t.status === 'COMPLETED')
    .reduce((sum, t) => sum + (t.netAmount || t.amount), 0)

  const filteredContracts = (tab: string) => {
    if (tab === 'active') return contracts.filter(c => c.status === 'ACTIVE' || c.status === 'PAUSED')
    if (tab === 'completed') return contracts.filter(c => c.status === 'COMPLETED' || c.status === 'CANCELLED')
    return contracts
  }

  const getProgress = (contract: Contract) => {
    const ms = contract.milestones || []
    if (ms.length === 0) return 0
    const completed = ms.filter(m => m.status === 'PAID' || m.status === 'APPROVED').length
    return Math.round((completed / ms.length) * 100)
  }

  const getOtherParty = (contract: Contract) => {
    if (user?.id === contract.client?.id) return contract.freelancer
    return contract.client
  }

  const ContractCard = ({ contract }: { contract: Contract }) => {
    const otherParty = getOtherParty(contract)
    const progress = getProgress(contract)
    const ms = contract.milestones || []
    const completedMs = ms.filter(m => m.status === 'PAID' || m.status === 'APPROVED').length

    return (
      <Card className="border-slate-200 hover:shadow-sm transition-shadow">
        <CardContent className="p-4">
          <div className="flex items-start gap-3">
            <Avatar className="h-11 w-11 border border-slate-200 shrink-0">
              <AvatarFallback className="bg-emerald-100 text-emerald-700 text-xs font-semibold">
                {otherParty?.displayName?.[0] || '?'}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <div
                  className="text-sm font-semibold text-slate-900 hover:text-emerald-600 cursor-pointer truncate"
                  onClick={() => navigate('contracts/detail', { id: contract.id })}
                >
                  {contract.title}
                </div>
                <ContractStatusBadge status={contract.status} />
              </div>
              <div className="text-xs text-slate-500 mt-0.5">
                {otherParty?.displayName || 'Anonymous'} &middot; {user?.id === contract.client?.id ? 'Freelancer' : 'Client'}
              </div>
              <div className="flex items-center gap-4 mt-2">
                <span className="text-sm font-medium text-emerald-600">
                  {contract.contractType === 'FIXED_PRICE'
                    ? formatCurrency(contract.totalAmount)
                    : `$${contract.hourlyRate || 0}/hr`}
                </span>
                {ms.length > 0 && (
                  <div className="flex items-center gap-2 text-xs text-slate-400">
                    <span>{completedMs}/{ms.length} milestones</span>
                    <Progress value={progress} className="h-1.5 w-20" />
                  </div>
                )}
              </div>
              <div className="flex items-center justify-between mt-3">
                <span className="text-xs text-slate-400 flex items-center gap-1">
                  <Clock className="w-3 h-3" />
                  {timeAgo(contract.updatedAt)}
                </span>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => navigate('contracts/detail', { id: contract.id })}
                  className="text-xs h-8 border-emerald-200 text-emerald-600 hover:bg-emerald-50"
                >
                  View
                  <ArrowRight className="w-3 h-3 ml-1" />
                </Button>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>
    )
  }

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-48" />
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[1, 2, 3, 4].map(i => <Skeleton key={i} className="h-20 rounded-xl" />)}
        </div>
        <Skeleton className="h-10 w-64" />
        {[1, 2, 3].map(i => <Skeleton key={i} className="h-28 rounded-xl" />)}
      </div>
    )
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">My Contracts</h1>
        <p className="text-slate-500 mt-1">View and manage all your contracts.</p>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <StatCard
          label="Total"
          value={contracts.length}
          icon={Briefcase}
          iconColor="bg-slate-100 text-slate-600"
        />
        <StatCard
          label="Active"
          value={activeCount}
          icon={Briefcase}
          iconColor="bg-emerald-100 text-emerald-600"
        />
        <StatCard
          label="Completed"
          value={completedCount}
          icon={Briefcase}
          iconColor="bg-green-100 text-green-600"
        />
        <StatCard
          label={user?.role === 'FREELANCER' ? 'Earned' : 'Spent'}
          value={user?.role === 'FREELANCER' ? totalEarned : totalSpent}
          format="currency"
          icon={DollarSign}
          iconColor="bg-amber-100 text-amber-600"
        />
      </div>

      {/* Tabs */}
      <Tabs defaultValue="active">
        <TabsList className="bg-slate-100">
          <TabsTrigger value="active" className="text-sm">Active ({activeCount})</TabsTrigger>
          <TabsTrigger value="completed" className="text-sm">Completed ({completedCount})</TabsTrigger>
          <TabsTrigger value="all" className="text-sm">All ({contracts.length})</TabsTrigger>
        </TabsList>

        {['active', 'completed', 'all'].map(tab => (
          <TabsContent key={tab} value={tab}>
            {filteredContracts(tab).length === 0 ? (
              <Card className="border-slate-200">
                <CardContent className="p-12">
                  <EmptyState
                    icon={Briefcase}
                    title={tab === 'all' ? 'No contracts yet' : `No ${tab} contracts`}
                    description={tab === 'all'
                      ? "You don't have any contracts yet. Start by applying to jobs or posting a project."
                      : `You don't have any ${tab} contracts at the moment.`}
                    actionLabel="Browse Jobs"
                    onAction={() => navigate('jobs')}
                  />
                </CardContent>
              </Card>
            ) : (
              <div className="space-y-3">
                {filteredContracts(tab).map(contract => (
                  <ContractCard key={contract.id} contract={contract} />
                ))}
              </div>
            )}
          </TabsContent>
        ))}
      </Tabs>
    </motion.div>
  )
}
