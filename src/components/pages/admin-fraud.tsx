'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Badge } from '@/components/ui/badge'
import { Skeleton } from '@/components/ui/skeleton'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { ScrollArea } from '@/components/ui/scroll-area'
import { Separator } from '@/components/ui/separator'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from '@/components/ui/dialog'
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import { StatCard } from '@/components/shared/stat-card'
import { EmptyState } from '@/components/shared/empty-state'
import { authFetch } from '@/lib/api-fetch'
import { cn } from '@/lib/utils'
import { timeAgo } from '@/lib/format'
import {
  ShieldAlert,
  Eye,
  XCircle,
  Ban,
  RefreshCw,
  ChevronLeft,
  ChevronRight,
  AlertTriangle,
  AlertOctagon,
  Info,
  User,
  Clock,
  FileWarning,
} from 'lucide-react'

// ─── Types ────────────────────────────────────────────────────────────────────

interface FraudAlertItem {
  id: string
  userId: string
  alertType: string
  severity: string
  description: string
  metadata: string
  status: string
  reviewedBy: string | null
  reviewedAt: string | null
  adminNote: string | null
  createdAt: string
  user: {
    id: string
    email: string
    role: string
    isSuspended: boolean
    createdAt: string
    freelancerProfile?: { displayName: string } | null
    clientProfile?: { displayName: string } | null
  }
}

interface FraudAlertDetail {
  alert: FraudAlertItem & {
    user: {
      id: string
      email: string
      role: string
      isSuspended: boolean
      isVerified: boolean
      createdAt: string
      freelancerProfile?: { displayName: string; rating: number; completedJobs: number } | null
      clientProfile?: { displayName: string; rating: number; totalSpent: number } | null
    }
  }
  userReportHistory: {
    reportsReceived: Array<{
      id: string
      reportType: string
      description: string
      status: string
      createdAt: string
      reporter: {
        freelancerProfile?: { displayName: string } | null
        clientProfile?: { displayName: string } | null
      }
    }>
    reportsFiledCount: number
  }
  totalFraudAlerts: number
}

// ─── Constants ────────────────────────────────────────────────────────────────

const SEVERITY_CONFIG: Record<string, { color: string; bgColor: string; icon: React.ElementType; label: string }> = {
  CRITICAL: { color: 'text-red-700', bgColor: 'bg-red-100 border-red-200', icon: AlertOctagon, label: 'Critical' },
  HIGH: { color: 'text-orange-700', bgColor: 'bg-orange-100 border-orange-200', icon: AlertTriangle, label: 'High' },
  MEDIUM: { color: 'text-amber-700', bgColor: 'bg-amber-100 border-amber-200', icon: Info, label: 'Medium' },
  LOW: { color: 'text-slate-700', bgColor: 'bg-slate-100 border-slate-200', icon: Info, label: 'Low' },
}

const STATUS_CONFIG: Record<string, { color: string; label: string }> = {
  PENDING: { color: 'bg-amber-100 text-amber-700 border-amber-200', label: 'Pending' },
  REVIEWED: { color: 'bg-blue-100 text-blue-700 border-blue-200', label: 'Reviewed' },
  DISMISSED: { color: 'bg-slate-100 text-slate-600 border-slate-200', label: 'Dismissed' },
  ACTION_TAKEN: { color: 'bg-red-100 text-red-700 border-red-200', label: 'Action Taken' },
}

const ALERT_TYPE_LABELS: Record<string, string> = {
  SUSPICIOUS_ACTIVITY: 'Suspicious Activity',
  MULTIPLE_REPORTS: 'Multiple Reports',
  RAPID_SIGNUP: 'Rapid Signup',
  PAYMENT_ANOMALY: 'Payment Anomaly',
  FAKE_PROFILE: 'Fake Profile',
  ACCOUNT_TAKEOVER: 'Account Takeover',
}

// ─── Component ────────────────────────────────────────────────────────────────

export function AdminFraud() {
  const [alerts, setAlerts] = useState<FraudAlertItem[]>([])
  const [loading, setLoading] = useState(true)
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [statusFilter, setStatusFilter] = useState<string>('PENDING')
  const [severityFilter, setSeverityFilter] = useState<string>('all')

  // Stats
  const [stats, setStats] = useState({ pending: 0, highPending: 0, criticalPending: 0 })

  // Detail dialog
  const [selectedAlert, setSelectedAlert] = useState<FraudAlertDetail | null>(null)
  const [detailLoading, setDetailLoading] = useState(false)
  const [detailOpen, setDetailOpen] = useState(false)

  // Dismiss dialog
  const [dismissOpen, setDismissOpen] = useState(false)
  const [dismissNote, setDismissNote] = useState('')
  const [dismissLoading, setDismissLoading] = useState(false)

  // Suspend dialog
  const [suspendOpen, setSuspendOpen] = useState(false)
  const [suspendNote, setSuspendNote] = useState('')
  const [suspendLoading, setSuspendLoading] = useState(false)

  const fetchAlerts = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', String(page))
      params.set('limit', '20')
      if (statusFilter !== 'all') params.set('status', statusFilter)
      if (severityFilter !== 'all') params.set('severity', severityFilter)

      const res = await authFetch(`/api/fraud-alerts?${params}`)
      if (res.ok) {
        const data = await res.json()
        setAlerts(data.alerts || [])
        setStats(data.stats || { pending: 0, highPending: 0, criticalPending: 0 })
        setTotalPages(data.pagination?.totalPages || 1)
      }
    } catch {
      // silently handle
    } finally {
      setLoading(false)
    }
  }, [page, statusFilter, severityFilter])

  // Auto-refresh every 30 seconds
  useEffect(() => {
    fetchAlerts()
    const interval = setInterval(() => { fetchAlerts() }, 30000)
    return () => clearInterval(interval)
  }, [fetchAlerts])

  // Reset to page 1 when filters change
  useEffect(() => { setPage(1) }, [statusFilter, severityFilter])

  const openDetail = async (alertId: string) => {
    setDetailLoading(true)
    setDetailOpen(true)
    try {
      const res = await authFetch(`/api/fraud-alerts/${alertId}`)
      if (res.ok) {
        const data = await res.json()
        setSelectedAlert(data)
      }
    } catch {
      // silently handle
    } finally {
      setDetailLoading(false)
    }
  }

  const handleDismiss = async () => {
    if (!selectedAlert) return
    setDismissLoading(true)
    try {
      const res = await authFetch(`/api/fraud-alerts/${selectedAlert.alert.id}`, {
        method: 'PATCH',
        body: JSON.stringify({ status: 'DISMISSED', adminNote: dismissNote || undefined }),
      })
      if (res.ok) {
        setDismissOpen(false)
        setDismissNote('')
        setDetailOpen(false)
        fetchAlerts()
      }
    } catch {
      // silently handle
    } finally {
      setDismissLoading(false)
    }
  }

  const handleMarkReviewed = async () => {
    if (!selectedAlert) return
    try {
      const res = await authFetch(`/api/fraud-alerts/${selectedAlert.alert.id}`, {
        method: 'PATCH',
        body: JSON.stringify({ status: 'REVIEWED' }),
      })
      if (res.ok) {
        setDetailOpen(false)
        fetchAlerts()
      }
    } catch {
      // silently handle
    }
  }

  const handleSuspend = async () => {
    if (!selectedAlert) return
    setSuspendLoading(true)
    try {
      const res = await authFetch(`/api/fraud-alerts/${selectedAlert.alert.id}/suspend`, {
        method: 'POST',
        body: JSON.stringify({ adminNote: suspendNote || undefined }),
      })
      if (res.ok) {
        setSuspendOpen(false)
        setSuspendNote('')
        setDetailOpen(false)
        fetchAlerts()
      }
    } catch {
      // silently handle
    } finally {
      setSuspendLoading(false)
    }
  }

  const getUserDisplayName = (user: FraudAlertItem['user']) => {
    return user.freelancerProfile?.displayName || user.clientProfile?.displayName || user.email
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-6"
    >
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 flex items-center gap-2">
            <ShieldAlert className="w-6 h-6 text-red-500" />
            Fraud Detection
          </h1>
          <p className="text-slate-500 mt-1">Monitor and manage fraud alerts across the platform</p>
        </div>
        <Button variant="outline" size="sm" onClick={() => { fetchAlerts() }}>
          <RefreshCw className="w-4 h-4 mr-2" />
          Refresh
        </Button>
      </div>

      {/* Stats Cards */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
        <StatCard
          label="Pending Alerts"
          value={stats.pending}
          icon={ShieldAlert}
          iconColor={stats.pending > 0 ? 'bg-amber-100 text-amber-600' : 'bg-slate-100 text-slate-500'}
        />
        <StatCard
          label="Critical"
          value={stats.criticalPending}
          icon={AlertOctagon}
          iconColor={stats.criticalPending > 0 ? 'bg-red-100 text-red-600' : 'bg-slate-100 text-slate-500'}
        />
        <StatCard
          label="High Severity"
          value={stats.highPending}
          icon={AlertTriangle}
          iconColor={stats.highPending > 0 ? 'bg-orange-100 text-orange-600' : 'bg-slate-100 text-slate-500'}
        />
        <StatCard
          label="Auto-Refresh"
          value={30}
          suffix="s"
          icon={Clock}
          iconColor="bg-emerald-100 text-emerald-600"
        />
      </div>

      {/* Filters */}
      <Card>
        <CardContent className="p-4">
          <div className="flex flex-col sm:flex-row gap-3">
            <Select value={statusFilter} onValueChange={setStatusFilter}>
              <SelectTrigger className="w-full sm:w-48">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Statuses</SelectItem>
                <SelectItem value="PENDING">Pending</SelectItem>
                <SelectItem value="REVIEWED">Reviewed</SelectItem>
                <SelectItem value="DISMISSED">Dismissed</SelectItem>
                <SelectItem value="ACTION_TAKEN">Action Taken</SelectItem>
              </SelectContent>
            </Select>
            <Select value={severityFilter} onValueChange={setSeverityFilter}>
              <SelectTrigger className="w-full sm:w-48">
                <SelectValue placeholder="Severity" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Severity</SelectItem>
                <SelectItem value="CRITICAL">Critical</SelectItem>
                <SelectItem value="HIGH">High</SelectItem>
                <SelectItem value="MEDIUM">Medium</SelectItem>
                <SelectItem value="LOW">Low</SelectItem>
              </SelectContent>
            </Select>
            {(statusFilter !== 'all' || severityFilter !== 'all') && (
              <Button
                variant="ghost"
                size="sm"
                onClick={() => { setStatusFilter('PENDING'); setSeverityFilter('all') }}
              >
                <XCircle className="w-4 h-4 mr-1" />
                Clear Filters
              </Button>
            )}
          </div>
        </CardContent>
      </Card>

      {/* Alerts List */}
      {loading ? (
        <div className="space-y-3">
          {[1, 2, 3].map(i => (
            <Skeleton key={i} className="h-24 rounded-xl" />
          ))}
        </div>
      ) : alerts.length === 0 ? (
        <EmptyState
          icon={ShieldAlert}
          title="No fraud alerts found"
          description={statusFilter !== 'all' || severityFilter !== 'all'
            ? 'No alerts match your current filters. Try adjusting them.'
            : 'Great news! No fraud alerts have been detected yet.'}
          actionLabel={statusFilter !== 'all' || severityFilter !== 'all' ? 'Clear Filters' : undefined}
          onAction={statusFilter !== 'all' || severityFilter !== 'all'
            ? () => { setStatusFilter('PENDING'); setSeverityFilter('all') }
            : undefined}
        />
      ) : (
        <>
          <div className="space-y-3">
            {alerts.map((alert, idx) => {
              const sev = SEVERITY_CONFIG[alert.severity] || SEVERITY_CONFIG.LOW
              const st = STATUS_CONFIG[alert.status] || STATUS_CONFIG.PENDING
              const SevIcon = sev.icon
              const metadata = (() => { try { return JSON.parse(alert.metadata) } catch { return {} } })()

              return (
                <motion.div
                  key={alert.id}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  transition={{ delay: idx * 0.03 }}
                  className={cn(
                    'bg-white rounded-xl border p-4 hover:shadow-md transition-all cursor-pointer',
                    alert.status === 'PENDING' ? 'border-l-4 border-l-red-400' : 'border-slate-200'
                  )}
                  onClick={() => openDetail(alert.id)}
                >
                  <div className="flex flex-col sm:flex-row sm:items-start gap-3">
                    {/* Severity icon */}
                    <div className={cn('w-10 h-10 rounded-lg flex items-center justify-center shrink-0', sev.bgColor)}>
                      <SevIcon className={cn('w-5 h-5', sev.color)} />
                    </div>

                    {/* Content */}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap mb-1">
                        <span className={cn(
                          'px-2 py-0.5 rounded-full text-[10px] font-semibold uppercase border',
                          sev.bgColor, sev.color
                        )}>
                          {sev.label}
                        </span>
                        <Badge variant="outline" className="text-[10px] border-slate-200 text-slate-600">
                          {ALERT_TYPE_LABELS[alert.alertType] || alert.alertType}
                        </Badge>
                        <Badge variant="outline" className={cn('text-[10px] border', st.color)}>
                          {st.label}
                        </Badge>
                        {alert.user.isSuspended && (
                          <Badge className="bg-red-500 text-white text-[10px]">Suspended</Badge>
                        )}
                      </div>

                      <p className="text-sm text-slate-800 font-medium mb-1">{alert.description}</p>

                      <div className="flex items-center gap-3 text-xs text-slate-500">
                        <span className="flex items-center gap-1">
                          <User className="w-3 h-3" />
                          {getUserDisplayName(alert.user)}
                        </span>
                        <span>{alert.user.role}</span>
                        <span>{timeAgo(alert.createdAt)}</span>
                      </div>

                      {/* Metadata summary */}
                      {metadata.recentReportCount && (
                        <div className="text-xs text-slate-400 mt-1">
                          {metadata.recentReportCount} reports in last {metadata.periodDays || 7} days
                        </div>
                      )}
                    </div>

                    {/* Actions */}
                    <div className="flex items-center gap-2 shrink-0">
                      <Button variant="ghost" size="sm" className="text-slate-500 hover:text-slate-700">
                        <Eye className="w-4 h-4" />
                      </Button>
                    </div>
                  </div>
                </motion.div>
              )
            })}
          </div>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-center gap-2">
              <Button
                variant="outline"
                size="sm"
                disabled={page <= 1}
                onClick={() => setPage(p => p - 1)}
              >
                <ChevronLeft className="w-4 h-4" />
                Previous
              </Button>
              <span className="text-sm text-slate-500">
                Page {page} of {totalPages}
              </span>
              <Button
                variant="outline"
                size="sm"
                disabled={page >= totalPages}
                onClick={() => setPage(p => p + 1)}
              >
                Next
                <ChevronRight className="w-4 h-4" />
              </Button>
            </div>
          )}
        </>
      )}

      {/* ─── Detail Dialog ─── */}
      <Dialog open={detailOpen} onOpenChange={setDetailOpen}>
        <DialogContent className="max-w-2xl max-h-[85vh] overflow-y-auto">
          {detailLoading ? (
            <div className="space-y-4 py-8">
              <Skeleton className="h-8 w-64" />
              <Skeleton className="h-4 w-48" />
              <Skeleton className="h-32" />
            </div>
          ) : selectedAlert ? (
            <>
              <DialogHeader>
                <DialogTitle className="flex items-center gap-2">
                  <FileWarning className="w-5 h-5 text-red-500" />
                  Fraud Alert Detail
                </DialogTitle>
                <DialogDescription>
                  Review and take action on this fraud alert
                </DialogDescription>
              </DialogHeader>

              <div className="space-y-4 mt-4">
                {/* Alert Info */}
                <div className="grid grid-cols-2 gap-3">
                  <div className="p-3 bg-slate-50 rounded-lg">
                    <div className="text-xs text-slate-500 mb-0.5">Severity</div>
                    <Badge className={cn('border', (SEVERITY_CONFIG[selectedAlert.alert.severity] || SEVERITY_CONFIG.LOW).bgColor, (SEVERITY_CONFIG[selectedAlert.alert.severity] || SEVERITY_CONFIG.LOW).color)}>
                      {selectedAlert.alert.severity}
                    </Badge>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-lg">
                    <div className="text-xs text-slate-500 mb-0.5">Status</div>
                    <Badge variant="outline" className={cn('border', (STATUS_CONFIG[selectedAlert.alert.status] || STATUS_CONFIG.PENDING).color)}>
                      {(STATUS_CONFIG[selectedAlert.alert.status] || STATUS_CONFIG.PENDING).label}
                    </Badge>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-lg">
                    <div className="text-xs text-slate-500 mb-0.5">Alert Type</div>
                    <div className="text-sm font-medium text-slate-900">
                      {ALERT_TYPE_LABELS[selectedAlert.alert.alertType] || selectedAlert.alert.alertType}
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-lg">
                    <div className="text-xs text-slate-500 mb-0.5">Created</div>
                    <div className="text-sm font-medium text-slate-900">{timeAgo(selectedAlert.alert.createdAt)}</div>
                  </div>
                </div>

                {/* Description */}
                <div className="p-3 bg-slate-50 rounded-lg">
                  <div className="text-xs text-slate-500 mb-1">Description</div>
                  <p className="text-sm text-slate-800">{selectedAlert.alert.description}</p>
                </div>

                {/* Admin Note */}
                {selectedAlert.alert.adminNote && (
                  <div className="p-3 bg-amber-50 rounded-lg border border-amber-200">
                    <div className="text-xs text-amber-600 mb-1">Admin Note</div>
                    <p className="text-sm text-amber-800">{selectedAlert.alert.adminNote}</p>
                  </div>
                )}

                <Separator />

                {/* User Info */}
                <div>
                  <h4 className="text-sm font-semibold text-slate-900 mb-2">User Information</h4>
                  <div className="grid grid-cols-2 gap-3">
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Display Name</div>
                      <div className="text-sm font-medium text-slate-900">
                        {selectedAlert.alert.user.freelancerProfile?.displayName || selectedAlert.alert.user.clientProfile?.displayName || 'N/A'}
                      </div>
                    </div>
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Email</div>
                      <div className="text-sm font-medium text-slate-900">{selectedAlert.alert.user.email}</div>
                    </div>
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Role</div>
                      <div className="text-sm font-medium text-slate-900">{selectedAlert.alert.user.role}</div>
                    </div>
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Account Age</div>
                      <div className="text-sm font-medium text-slate-900">{timeAgo(selectedAlert.alert.user.createdAt)}</div>
                    </div>
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Total Fraud Alerts</div>
                      <div className="text-sm font-bold text-red-600">{selectedAlert.totalFraudAlerts}</div>
                    </div>
                    <div className="p-3 bg-slate-50 rounded-lg">
                      <div className="text-xs text-slate-500 mb-0.5">Status</div>
                      {selectedAlert.alert.user.isSuspended ? (
                        <Badge className="bg-red-500 text-white text-xs">Suspended</Badge>
                      ) : selectedAlert.alert.user.isVerified ? (
                        <Badge className="bg-emerald-100 text-emerald-700 text-xs">Verified</Badge>
                      ) : (
                        <Badge variant="outline" className="text-xs">Unverified</Badge>
                      )}
                    </div>
                  </div>
                </div>

                {/* Report History */}
                <div>
                  <h4 className="text-sm font-semibold text-slate-900 mb-2">
                    Report History ({selectedAlert.userReportHistory.reportsReceived.length} received, {selectedAlert.userReportHistory.reportsFiledCount} filed)
                  </h4>
                  {selectedAlert.userReportHistory.reportsReceived.length === 0 ? (
                    <div className="text-sm text-slate-500 p-3 bg-slate-50 rounded-lg">No reports received</div>
                  ) : (
                    <ScrollArea className="max-h-48">
                      <div className="space-y-2">
                        {selectedAlert.userReportHistory.reportsReceived.map(report => (
                          <div key={report.id} className="flex items-start gap-2 p-2 bg-slate-50 rounded-lg text-sm">
                            <Badge variant="outline" className="text-[10px] border-slate-200 shrink-0 mt-0.5">
                              {report.reportType}
                            </Badge>
                            <div className="flex-1 min-w-0">
                              <p className="text-slate-700 truncate">{report.description}</p>
                              <p className="text-xs text-slate-400">
                                by {report.reporter.freelancerProfile?.displayName || report.reporter.clientProfile?.displayName || 'Unknown'} · {timeAgo(report.createdAt)}
                              </p>
                            </div>
                            <Badge variant="outline" className={cn('text-[10px] border shrink-0', (STATUS_CONFIG[report.status] || STATUS_CONFIG.PENDING).color)}>
                              {(STATUS_CONFIG[report.status] || STATUS_CONFIG.PENDING).label}
                            </Badge>
                          </div>
                        ))}
                      </div>
                    </ScrollArea>
                  )}
                </div>

                {/* Actions */}
                {selectedAlert.alert.status === 'PENDING' && (
                  <>
                    <Separator />
                    <div className="flex flex-wrap gap-2">
                      <Button variant="outline" size="sm" onClick={handleMarkReviewed}>
                        <Eye className="w-4 h-4 mr-1" />
                        Mark Reviewed
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        className="text-slate-500 hover:text-slate-700"
                        onClick={() => { setDismissNote(''); setDismissOpen(true) }}
                      >
                        <XCircle className="w-4 h-4 mr-1" />
                        Dismiss
                      </Button>
                      <Button
                        variant="destructive"
                        size="sm"
                        disabled={selectedAlert.alert.user.isSuspended}
                        onClick={() => { setSuspendNote(''); setSuspendOpen(true) }}
                      >
                        <Ban className="w-4 h-4 mr-1" />
                        {selectedAlert.alert.user.isSuspended ? 'Already Suspended' : 'Suspend User'}
                      </Button>
                    </div>
                  </>
                )}
              </div>
            </>
          ) : null}
        </DialogContent>
      </Dialog>

      {/* ─── Dismiss Dialog ─── */}
      <AlertDialog open={dismissOpen} onOpenChange={setDismissOpen}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Dismiss Fraud Alert</AlertDialogTitle>
            <AlertDialogDescription>
              Are you sure you want to dismiss this alert? This will mark it as not requiring further action.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <div className="py-2">
            <Textarea
              placeholder="Optional: Add a note explaining why this alert is being dismissed..."
              value={dismissNote}
              onChange={e => setDismissNote(e.target.value)}
              rows={3}
            />
          </div>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={dismissLoading}>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={handleDismiss}
              disabled={dismissLoading}
              className="bg-slate-600 hover:bg-slate-700"
            >
              {dismissLoading ? 'Dismissing...' : 'Dismiss Alert'}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>

      {/* ─── Suspend Dialog ─── */}
      <AlertDialog open={suspendOpen} onOpenChange={setSuspendOpen}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle className="text-red-600">Suspend User Account</AlertDialogTitle>
            <AlertDialogDescription>
              This will immediately suspend the user&apos;s account. The user will not be able to log in or perform any actions. This action should be taken carefully.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <div className="py-2">
            <Textarea
              placeholder="Reason for suspension (will be visible to the user)..."
              value={suspendNote}
              onChange={e => setSuspendNote(e.target.value)}
              rows={3}
            />
          </div>
          <AlertDialogFooter>
            <AlertDialogCancel disabled={suspendLoading}>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={handleSuspend}
              disabled={suspendLoading}
              className="bg-red-600 hover:bg-red-700"
            >
              {suspendLoading ? 'Suspending...' : 'Confirm Suspension'}
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </motion.div>
  )
}
