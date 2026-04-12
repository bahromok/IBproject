'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Card, CardContent } from '@/components/ui/card'
import { Skeleton } from '@/components/ui/skeleton'
import { Textarea } from '@/components/ui/textarea'
import { Label } from '@/components/ui/label'
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import {
  Scale,
  ChevronLeft,
  ChevronRight,
  CheckCircle2,
  AlertTriangle,
  Eye,
  DollarSign,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { formatCurrency } from '@/lib/format'
import { authFetch } from '@/lib/api-fetch'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { toast } from 'sonner'

interface Dispute {
  id: string
  title: string
  status: string
  totalAmount: number
  contractType: string
  createdAt: string
  updatedAt: string
  job: { id: string; title: string; category: string } | null
  client: { id: string; clientProfile?: { displayName: string } | null } | null
  freelancer: { id: string; freelancerProfile?: { displayName: string } | null } | null
  milestones: { id: string; title: string; amount: number; status: string }[]
  _count?: { transactions: number }
}

export function AdminDisputes() {
  const { navigate } = useNavigationStore()
  const [disputes, setDisputes] = useState<Dispute[]>([])
  const [loading, setLoading] = useState(true)
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [resolutionDialogOpen, setResolutionDialogOpen] = useState(false)
  const [selectedDispute, setSelectedDispute] = useState<Dispute | null>(null)
  const [resolution, setResolution] = useState('release')
  const [resolutionNotes, setResolutionNotes] = useState('')
  const [actionLoading, setActionLoading] = useState(false)

  const fetchDisputes = useCallback(async () => {
    setLoading(true)
    try {
      const res = await authFetch(`/api/admin/disputes?page=${page}&limit=20`)
      if (res.ok) {
        const data = await res.json()
        setDisputes(data.disputes || [])
        setTotalPages(data.pagination?.totalPages || 1)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [page])

  useEffect(() => { fetchDisputes() }, [fetchDisputes])

  const handleResolve = async () => {
    if (!selectedDispute) return
    setActionLoading(true)

    try {
      const res = await authFetch('/api/admin/disputes', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contractId: selectedDispute.id,
          resolution,
          notes: resolutionNotes,
        }),
      })

      if (res.ok) {
        toast.success(`Dispute resolved: ${resolution === 'release' ? 'Funds released to freelancer' : 'Refund issued to client'}`)
        fetchDisputes()
        setResolutionDialogOpen(false)
        setSelectedDispute(null)
      } else {
        const data = await res.json()
        toast.error(data.error || 'Failed to resolve dispute')
      }
    } catch {
      toast.error('An error occurred')
    } finally {
      setActionLoading(false)
    }
  }

  const openResolutionDialog = (dispute: Dispute) => {
    setSelectedDispute(dispute)
    setResolution('release')
    setResolutionNotes('')
    setResolutionDialogOpen(true)
  }

  const totalDisputedAmount = disputes.reduce((sum, d) => sum + (d.totalAmount || 0), 0)

  if (loading) {
    return (
      <div className="space-y-6">
        <div>
          <Skeleton className="h-8 w-48 mb-2" />
          <Skeleton className="h-4 w-64" />
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          {[1, 2, 3].map(i => <Skeleton key={i} className="h-24 rounded-xl" />)}
        </div>
        {[1, 2, 3].map(i => <Skeleton key={i} className="h-40 rounded-xl" />)}
      </div>
    )
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-6"
    >
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Dispute Resolution</h1>
        <p className="text-sm text-slate-500 mt-0.5">Manage and resolve contract disputes</p>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard
          label="Open Disputes"
          value={disputes.length}
          icon={AlertTriangle}
          iconColor={disputes.length > 0 ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-600'}
        />
        <StatCard
          label="Total Disputed Amount"
          value={totalDisputedAmount}
          format="currency"
          icon={DollarSign}
          iconColor="bg-amber-100 text-amber-600"
        />
        <StatCard
          label="Contracts with Milestones"
          value={disputes.filter(d => d.milestones && d.milestones.length > 0).length}
          icon={Scale}
          iconColor="bg-violet-100 text-violet-600"
        />
      </div>

      {/* Disputes List */}
      {disputes.length === 0 ? (
        <Card className="border-slate-200">
          <CardContent className="p-12">
            <EmptyState
              icon={Scale}
              title="No open disputes"
              description="Great news! There are no active disputes to resolve."
            />
          </CardContent>
        </Card>
      ) : (
        <div className="space-y-4">
          {disputes.map((dispute) => {
            const clientName = dispute.client?.clientProfile?.displayName || 'Client'
            const freelancerName = dispute.freelancer?.freelancerProfile?.displayName || 'Freelancer'

            return (
              <Card key={dispute.id} className="hover:shadow-md transition-shadow">
                <CardContent className="p-5">
                  <div className="flex flex-col lg:flex-row lg:items-center gap-4">
                    {/* Info */}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 mb-2">
                        <Badge variant="secondary" className="bg-red-100 text-red-700 text-xs">
                          <Scale className="w-3 h-3 mr-0.5" /> Disputed
                        </Badge>
                        <span className="text-xs text-slate-400">
                          Opened {new Date(dispute.updatedAt || dispute.createdAt).toLocaleDateString()}
                        </span>
                      </div>
                      <h3 className="text-base font-semibold text-slate-900 mb-1">{dispute.title}</h3>
                      <div className="flex items-center gap-2 text-sm text-slate-600">
                        <span className="font-medium text-emerald-700">{clientName}</span>
                        <span className="text-slate-400">vs</span>
                        <span className="font-medium text-violet-700">{freelancerName}</span>
                      </div>

                      {/* Milestones */}
                      {dispute.milestones && dispute.milestones.length > 0 && (
                        <div className="mt-3 flex flex-wrap gap-1.5">
                          {dispute.milestones.map(ms => (
                            <Badge key={ms.id} variant="outline" className="text-xs">
                              {ms.title} - {formatCurrency(ms.amount)} ({ms.status})
                            </Badge>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Amount & Actions */}
                    <div className="flex items-center gap-4 lg:flex-col lg:items-end shrink-0">
                      <div className="text-right">
                        <div className="text-2xl font-bold text-slate-900">
                          {formatCurrency(dispute.totalAmount)}
                        </div>
                        <div className="text-xs text-slate-500">Disputed Amount</div>
                      </div>
                      <div className="flex items-center gap-2">
                        <Button
                          variant="outline"
                          size="sm"
                          className="h-8"
                          onClick={() => navigate('contracts', { id: dispute.id })}
                        >
                          <Eye className="w-3.5 h-3.5 mr-1" /> View Contract
                        </Button>
                        <Button
                          size="sm"
                          className="h-8 bg-emerald-500 hover:bg-emerald-600"
                          onClick={() => openResolutionDialog(dispute)}
                        >
                          <Scale className="w-3.5 h-3.5 mr-1" /> Resolve
                        </Button>
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>
            )
          })}
        </div>
      )}

      {/* Pagination */}
      {totalPages > 1 && (
        <div className="flex items-center justify-between">
          <span className="text-sm text-slate-500">Page {page} of {totalPages}</span>
          <div className="flex items-center gap-1">
            <Button variant="outline" size="icon" className="h-8 w-8" onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page <= 1}>
              <ChevronLeft className="w-4 h-4" />
            </Button>
            <Button variant="outline" size="icon" className="h-8 w-8" onClick={() => setPage(p => Math.min(totalPages, p + 1))} disabled={page >= totalPages}>
              <ChevronRight className="w-4 h-4" />
            </Button>
          </div>
        </div>
      )}

      {/* Resolution Dialog */}
      <Dialog open={resolutionDialogOpen} onOpenChange={setResolutionDialogOpen}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Resolve Dispute</DialogTitle>
          </DialogHeader>
          {selectedDispute && (
            <div className="space-y-5 mt-2">
              <div className="bg-slate-50 p-3 rounded-lg">
                <p className="text-sm font-medium text-slate-900">{selectedDispute.title}</p>
                <p className="text-sm text-slate-500 mt-1">
                  Disputed amount: <span className="font-semibold">{formatCurrency(selectedDispute.totalAmount)}</span>
                </p>
              </div>

              <RadioGroup value={resolution} onValueChange={setResolution} className="space-y-2">
                <label
                  className={cn(
                    'flex items-center gap-3 p-3 rounded-lg border-2 cursor-pointer transition-all',
                    resolution === 'release'
                      ? 'border-emerald-500 bg-emerald-50/50'
                      : 'border-slate-200 hover:border-slate-300'
                  )}
                >
                  <RadioGroupItem value="release" />
                  <div>
                    <p className="text-sm font-medium text-slate-900">Release to Freelancer</p>
                    <p className="text-xs text-slate-500">Release full amount ({formatCurrency(selectedDispute.totalAmount)})</p>
                  </div>
                </label>
                <label
                  className={cn(
                    'flex items-center gap-3 p-3 rounded-lg border-2 cursor-pointer transition-all',
                    resolution === 'refund'
                      ? 'border-amber-500 bg-amber-50/50'
                      : 'border-slate-200 hover:border-slate-300'
                  )}
                >
                  <RadioGroupItem value="refund" />
                  <div>
                    <p className="text-sm font-medium text-slate-900">Refund to Client</p>
                    <p className="text-xs text-slate-500">Refund full amount ({formatCurrency(selectedDispute.totalAmount)})</p>
                  </div>
                </label>
              </RadioGroup>

              <div className="space-y-2">
                <Label htmlFor="notes">Resolution Notes *</Label>
                <Textarea
                  id="notes"
                  placeholder="Describe the resolution decision..."
                  value={resolutionNotes}
                  onChange={e => setResolutionNotes(e.target.value)}
                  rows={3}
                />
              </div>

              <div className="flex justify-end gap-2 pt-1">
                <Button variant="outline" onClick={() => setResolutionDialogOpen(false)}>Cancel</Button>
                <Button
                  onClick={handleResolve}
                  disabled={actionLoading || !resolutionNotes.trim()}
                  className={resolution === 'release' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-amber-600 hover:bg-amber-700'}
                >
                  {actionLoading ? 'Processing...' : 'Apply Resolution'}
                </Button>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </motion.div>
  )
}
