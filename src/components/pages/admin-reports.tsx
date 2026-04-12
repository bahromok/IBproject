'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Card, CardContent } from '@/components/ui/card'
import { Skeleton } from '@/components/ui/skeleton'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Textarea } from '@/components/ui/textarea'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import {
  AlertTriangle,
  Ban,
  CheckCircle2,
  Clock,
  XCircle,
  Search,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { EmptyState } from '@/components/shared/empty-state'
import { authFetch } from '@/lib/api-fetch'
import { toast } from 'sonner'

const STATUS_CONFIG: Record<string, { label: string; color: string; icon: React.ElementType }> = {
  PENDING: { label: 'Pending', color: 'bg-amber-100 text-amber-700', icon: Clock },
  UNDER_REVIEW: { label: 'Under Review', color: 'bg-emerald-100 text-emerald-700', icon: Search },
  RESOLVED: { label: 'Resolved', color: 'bg-emerald-100 text-emerald-700', icon: CheckCircle2 },
  DISMISSED: { label: 'Dismissed', color: 'bg-slate-100 text-slate-600', icon: XCircle },
}

const REPORT_TYPE_LABELS: Record<string, string> = {
  SCAM: 'Scam/Fraud',
  HARASSMENT: 'Harassment',
  SPAM: 'Spam',
  INAPPROPRIATE_CONTENT: 'Inappropriate',
  PAYMENT_DISPUTE: 'Payment Dispute',
  OTHER: 'Other',
}

interface Report {
  id: string
  reporterId: string
  reportedUserId: string
  reportType: string
  description: string
  evidence: string
  status: string
  adminNote: string | null
  createdAt: string
  reporter?: {
    id: string
    email: string
    clientProfile?: { displayName: string } | null
    freelancerProfile?: { displayName: string } | null
  }
}

export function AdminReports() {
  const { navigate } = useNavigationStore()
  const [reports, setReports] = useState<Report[]>([])
  const [loading, setLoading] = useState(true)
  const [activeTab, setActiveTab] = useState('all')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [selectedReport, setSelectedReport] = useState<Report | null>(null)
  const [detailDialogOpen, setDetailDialogOpen] = useState(false)
  const [noteDialogOpen, setNoteDialogOpen] = useState(false)
  const [noteAction, setNoteAction] = useState<'resolve' | 'dismiss'>('resolve')
  const [adminNote, setAdminNote] = useState('')
  const [actionLoading, setActionLoading] = useState(false)

  const fetchReports = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', page.toString())
      params.set('limit', '20')
      if (activeTab !== 'all') params.set('status', activeTab)

      const res = await authFetch(`/api/admin/reports?${params.toString()}`)
      if (res.ok) {
        const data = await res.json()
        setReports(data.reports || [])
        setTotalPages(data.pagination?.totalPages || 1)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [page, activeTab])

  useEffect(() => { fetchReports() }, [fetchReports])

  const handleStatusChange = async (reportId: string, status: string, note?: string) => {
    setActionLoading(true)
    try {
      const res = await authFetch('/api/admin/reports', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reportId, status, adminNote: note }),
      })
      if (res.ok) {
        toast.success(`Report ${status === 'RESOLVED' ? 'resolved' : status === 'DISMISSED' ? 'dismissed' : 'updated'}`)
        fetchReports()
        setNoteDialogOpen(false)
        setDetailDialogOpen(false)
      } else {
        toast.error('Failed to update report')
      }
    } catch {
      toast.error('An error occurred')
    } finally {
      setActionLoading(false)
    }
  }

  const openNoteDialog = (report: Report, action: 'resolve' | 'dismiss') => {
    setSelectedReport(report)
    setNoteAction(action)
    setAdminNote('')
    setNoteDialogOpen(true)
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-4"
    >
      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Report Management</h1>
        <p className="text-sm text-slate-500 mt-0.5">Review and handle user reports</p>
      </div>

      {/* Tabs & Table */}
      <Card>
        <CardContent className="p-0">
          <Tabs value={activeTab} onValueChange={v => { setActiveTab(v); setPage(1) }}>
            <div className="border-b border-slate-200 px-4">
              <TabsList className="bg-transparent h-auto p-0 gap-0">
                {['all', 'PENDING', 'UNDER_REVIEW', 'RESOLVED', 'DISMISSED'].map(tab => (
                  <TabsTrigger
                    key={tab}
                    value={tab}
                    className={cn(
                      'px-4 py-3 rounded-none border-b-2 border-transparent data-[state=active]:border-emerald-500 data-[state=active]:bg-transparent data-[state=active]:shadow-none text-sm font-medium text-slate-500 data-[state=active]:text-slate-900 transition-colors'
                    )}
                  >
                    {tab === 'all' ? 'All' : STATUS_CONFIG[tab]?.label || tab}
                  </TabsTrigger>
                ))}
              </TabsList>
            </div>

            <TabsContent value={activeTab} className="m-0">
              {loading ? (
                <div className="space-y-3 p-4">
                  {[1, 2, 3, 4, 5].map(i => (
                    <Skeleton key={i} className="h-20 w-full rounded-lg" />
                  ))}
                </div>
              ) : reports.length === 0 ? (
                <div className="py-12">
                  <EmptyState
                    icon={AlertTriangle}
                    title="No reports found"
                    description={activeTab !== 'all'
                      ? `No ${STATUS_CONFIG[activeTab]?.label?.toLowerCase() || activeTab} reports.`
                      : 'No reports have been submitted yet.'}
                    actionLabel={activeTab !== 'all' ? 'View All Reports' : undefined}
                    onAction={activeTab !== 'all' ? () => setActiveTab('all') : undefined}
                  />
                </div>
              ) : (
                <div className="divide-y divide-slate-100">
                  {reports.map((report) => {
                    const reporterName = report.reporter?.clientProfile?.displayName
                      || report.reporter?.freelancerProfile?.displayName
                      || 'Unknown'
                    const config = STATUS_CONFIG[report.status] || STATUS_CONFIG.PENDING
                    const StatusIcon = config.icon

                    return (
                      <div
                        key={report.id}
                        className="flex items-start gap-3 p-4 hover:bg-slate-50/50 transition-colors cursor-pointer"
                        onClick={() => { setSelectedReport(report); setDetailDialogOpen(true) }}
                      >
                        <div className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center shrink-0 mt-0.5">
                          <AlertTriangle className="w-5 h-5 text-slate-500" />
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-1 flex-wrap">
                            <Badge variant="secondary" className={cn('text-xs', config.color)}>
                              <StatusIcon className="w-3 h-3 mr-0.5" />
                              {config.label}
                            </Badge>
                            <Badge variant="outline" className="text-xs">
                              {REPORT_TYPE_LABELS[report.reportType] || report.reportType}
                            </Badge>
                            <span className="text-xs text-slate-400">
                              {new Date(report.createdAt).toLocaleDateString()}
                            </span>
                          </div>
                          <p className="text-sm font-medium text-slate-900 mb-1">
                            Reporter: {reporterName}
                          </p>
                          <p className="text-sm text-slate-500 line-clamp-2">{report.description}</p>
                          {report.adminNote && (
                            <div className="mt-2 p-2 bg-amber-50 rounded-md border border-amber-100">
                              <p className="text-xs text-amber-800">
                                <strong>Admin Note:</strong> {report.adminNote}
                              </p>
                            </div>
                          )}
                        </div>
                        <div className="flex items-center gap-1 shrink-0">
                          {report.status === 'PENDING' && (
                            <Button
                              size="sm"
                              variant="outline"
                              className="h-7 text-xs"
                              onClick={(e) => {
                                e.stopPropagation()
                                handleStatusChange(report.id, 'UNDER_REVIEW')
                              }}
                            >
                              Start Review
                            </Button>
                          )}
                          {report.status === 'UNDER_REVIEW' && (
                            <>
                              <Button
                                size="sm"
                                variant="outline"
                                className="h-7 text-xs text-emerald-600"
                                onClick={(e) => {
                                  e.stopPropagation()
                                  openNoteDialog(report, 'resolve')
                                }}
                              >
                                Resolve
                              </Button>
                              <Button
                                size="sm"
                                variant="outline"
                                className="h-7 text-xs text-slate-500"
                                onClick={(e) => {
                                  e.stopPropagation()
                                  openNoteDialog(report, 'dismiss')
                                }}
                              >
                                Dismiss
                              </Button>
                            </>
                          )}
                        </div>
                      </div>
                    )
                  })}
                </div>
              )}
            </TabsContent>
          </Tabs>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100">
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
        </CardContent>
      </Card>

      {/* Detail Dialog */}
      <Dialog open={detailDialogOpen} onOpenChange={setDetailDialogOpen}>
        <DialogContent className="max-w-lg">
          <DialogHeader>
            <DialogTitle>Report Details</DialogTitle>
          </DialogHeader>
          {selectedReport && (
            <div className="space-y-4 mt-2">
              <div className="grid grid-cols-2 gap-3 text-sm">
                <div>
                  <span className="text-slate-500">Reporter:</span>
                  <p className="font-medium text-slate-900">
                    {selectedReport.reporter?.clientProfile?.displayName
                      || selectedReport.reporter?.freelancerProfile?.displayName
                      || 'Unknown'}
                  </p>
                </div>
                <div>
                  <span className="text-slate-500">Reported User:</span>
                  <p className="font-medium text-slate-900">{selectedReport.reportedUserId?.slice(0, 8)}...</p>
                </div>
                <div>
                  <span className="text-slate-500">Type:</span>
                  <p className="font-medium">{REPORT_TYPE_LABELS[selectedReport.reportType] || selectedReport.reportType}</p>
                </div>
                <div>
                  <span className="text-slate-500">Date:</span>
                  <p className="font-medium">{new Date(selectedReport.createdAt).toLocaleString()}</p>
                </div>
              </div>
              <div>
                <span className="text-sm text-slate-500">Description:</span>
                <p className="text-sm text-slate-900 bg-slate-50 p-3 rounded-lg mt-1">{selectedReport.description}</p>
              </div>
              {selectedReport.adminNote && (
                <div>
                  <span className="text-sm text-slate-500">Admin Note:</span>
                  <p className="text-sm text-slate-900 bg-amber-50 border border-amber-100 p-3 rounded-lg mt-1">{selectedReport.adminNote}</p>
                </div>
              )}
              <div className="flex flex-wrap gap-2 pt-2">
                {selectedReport.status === 'PENDING' && (
                  <Button size="sm" onClick={() => handleStatusChange(selectedReport.id, 'UNDER_REVIEW')}>
                    Start Review
                  </Button>
                )}
                {selectedReport.status === 'UNDER_REVIEW' && (
                  <>
                    <Button size="sm" variant="outline" className="text-emerald-600" onClick={() => openNoteDialog(selectedReport, 'resolve')}>
                      Resolve
                    </Button>
                    <Button size="sm" variant="outline" className="text-slate-500" onClick={() => openNoteDialog(selectedReport, 'dismiss')}>
                      Dismiss
                    </Button>
                  </>
                )}
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>

      {/* Note Dialog */}
      <Dialog open={noteDialogOpen} onOpenChange={setNoteDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>
              {noteAction === 'resolve' ? 'Resolve Report' : 'Dismiss Report'}
            </DialogTitle>
          </DialogHeader>
          <div className="space-y-4 mt-4">
            <Textarea
              placeholder={noteAction === 'resolve' ? 'Resolution notes (optional)...' : 'Reason for dismissal (optional)...'}
              value={adminNote}
              onChange={e => setAdminNote(e.target.value)}
              rows={3}
            />
            <div className="flex justify-end gap-2">
              <Button variant="outline" onClick={() => setNoteDialogOpen(false)}>Cancel</Button>
              <Button
                onClick={() => selectedReport && handleStatusChange(
                  selectedReport.id,
                  noteAction === 'resolve' ? 'RESOLVED' : 'DISMISSED',
                  adminNote
                )}
                disabled={actionLoading}
                className={noteAction === 'resolve' ? 'bg-emerald-600 hover:bg-emerald-700' : ''}
              >
                {actionLoading ? 'Processing...' : noteAction === 'resolve' ? 'Resolve' : 'Dismiss'}
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </motion.div>
  )
}
