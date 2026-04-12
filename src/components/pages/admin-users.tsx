'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { Card, CardContent } from '@/components/ui/card'
import { Avatar, AvatarFallback, AvatarImage } from '@/components/ui/avatar'
import { Skeleton } from '@/components/ui/skeleton'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Textarea } from '@/components/ui/textarea'
import { Checkbox } from '@/components/ui/checkbox'
import {
  Search,
  Users,
  Shield,
  ShieldOff,
  Ban,
  CheckCircle2,
  Eye,
  MoreVertical,
  ChevronLeft,
  ChevronRight,
} from 'lucide-react'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { cn } from '@/lib/utils'
import { EmptyState } from '@/components/shared/empty-state'
import { authFetch } from '@/lib/api-fetch'
import { toast } from 'sonner'

interface AdminUser {
  id: string
  email: string
  name: string | null
  role: string
  isVerified: boolean
  isSuspended: boolean
  createdAt: string
  clientProfile?: { displayName: string; id: string } | null
  freelancerProfile?: { displayName: string; id: string; rating: number; completedJobs: number } | null
  _count?: { jobs: number; contractsAsClient: number; contractsAsFreelancer: number }
}

export function AdminUsers() {
  const { navigate } = useNavigationStore()
  const [users, setUsers] = useState<AdminUser[]>([])
  const [loading, setLoading] = useState(true)
  const [search, setSearch] = useState('')
  const [roleFilter, setRoleFilter] = useState('all')
  const [statusFilter, setStatusFilter] = useState('all')
  const [page, setPage] = useState(1)
  const [totalPages, setTotalPages] = useState(1)
  const [selectedUsers, setSelectedUsers] = useState<Set<string>>(new Set())
  const [suspendDialogOpen, setSuspendDialogOpen] = useState(false)
  const [suspendTarget, setSuspendTarget] = useState<string | null>(null)
  const [suspendReason, setSuspendReason] = useState('')
  const [actionLoading, setActionLoading] = useState(false)

  const fetchUsers = useCallback(async () => {
    setLoading(true)
    try {
      const params = new URLSearchParams()
      params.set('page', page.toString())
      params.set('limit', '25')
      if (search) params.set('search', search)
      if (roleFilter !== 'all') params.set('role', roleFilter)
      if (statusFilter === 'verified') params.set('isVerified', 'true')
      if (statusFilter === 'suspended') params.set('isSuspended', 'true')

      const res = await authFetch(`/api/admin/users?${params.toString()}`)
      if (res.ok) {
        const data = await res.json()
        setUsers(data.users || [])
        setTotalPages(data.pagination?.totalPages || 1)
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [page, search, roleFilter, statusFilter])

  useEffect(() => { fetchUsers() }, [fetchUsers])

  const handleAction = async (userId: string, action: string, reason?: string) => {
    setActionLoading(true)
    try {
      const res = await authFetch('/api/admin/users', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ userId, action, reason }),
      })
      if (res.ok) {
        toast.success(`Action "${action}" applied successfully`)
        fetchUsers()
      } else {
        const data = await res.json()
        toast.error(data.error || 'Action failed')
      }
    } catch {
      toast.error('An error occurred')
    } finally {
      setActionLoading(false)
      setSuspendDialogOpen(false)
      setSuspendReason('')
    }
  }

  const handleBulkAction = async (action: string) => {
    if (selectedUsers.size === 0) return
    setActionLoading(true)
    try {
      const promises = Array.from(selectedUsers).map(userId =>
        authFetch('/api/admin/users', {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ userId, action }),
        })
      )
      await Promise.all(promises)
      toast.success(`${selectedUsers.size} users updated`)
      setSelectedUsers(new Set())
      fetchUsers()
    } catch {
      toast.error('Bulk action failed')
    } finally {
      setActionLoading(false)
    }
  }

  const toggleSelect = (userId: string) => {
    const next = new Set(selectedUsers)
    if (next.has(userId)) next.delete(userId)
    else next.add(userId)
    setSelectedUsers(next)
  }

  const toggleSelectAll = () => {
    if (selectedUsers.size === users.length) {
      setSelectedUsers(new Set())
    } else {
      setSelectedUsers(new Set(users.map(u => u.id)))
    }
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-4"
    >
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">User Management</h1>
          <p className="text-sm text-slate-500 mt-0.5">Manage and moderate platform users</p>
        </div>
        {selectedUsers.size > 0 && (
          <div className="flex items-center gap-2">
            <span className="text-sm text-slate-500">{selectedUsers.size} selected</span>
            <Button
              size="sm"
              variant="outline"
              onClick={() => handleBulkAction('suspend')}
              disabled={actionLoading}
              className="text-red-600 border-red-200 hover:bg-red-50"
            >
              <Ban className="w-3.5 h-3.5 mr-1" /> Bulk Suspend
            </Button>
            <Button
              size="sm"
              variant="outline"
              onClick={() => handleBulkAction('verify')}
              disabled={actionLoading}
              className="text-emerald-600 border-emerald-200 hover:bg-emerald-50"
            >
              <CheckCircle2 className="w-3.5 h-3.5 mr-1" /> Bulk Verify
            </Button>
          </div>
        )}
      </div>

      {/* Filters */}
      <Card>
        <CardContent className="p-4">
          <div className="flex flex-col sm:flex-row gap-3">
            <div className="relative flex-1">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <Input
                placeholder="Search by display name or email..."
                className="pl-9 h-9"
                value={search}
                onChange={e => { setSearch(e.target.value); setPage(1) }}
              />
            </div>
            <Select value={roleFilter} onValueChange={v => { setRoleFilter(v); setPage(1) }}>
              <SelectTrigger className="h-9 w-full sm:w-40">
                <SelectValue placeholder="Role" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Roles</SelectItem>
                <SelectItem value="CLIENT">Clients</SelectItem>
                <SelectItem value="FREELANCER">Freelancers</SelectItem>
                <SelectItem value="ADMIN">Admins</SelectItem>
              </SelectContent>
            </Select>
            <Select value={statusFilter} onValueChange={v => { setStatusFilter(v); setPage(1) }}>
              <SelectTrigger className="h-9 w-full sm:w-40">
                <SelectValue placeholder="Status" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All Status</SelectItem>
                <SelectItem value="active">Active</SelectItem>
                <SelectItem value="verified">Verified</SelectItem>
                <SelectItem value="suspended">Suspended</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </CardContent>
      </Card>

      {/* Table */}
      <Card>
        <CardContent className="p-0">
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead className="w-10 pl-4">
                    <Checkbox
                      checked={users.length > 0 && selectedUsers.size === users.length}
                      onCheckedChange={toggleSelectAll}
                    />
                  </TableHead>
                  <TableHead>User</TableHead>
                  <TableHead className="hidden md:table-cell">Email</TableHead>
                  <TableHead>Role</TableHead>
                  <TableHead>Status</TableHead>
                  <TableHead className="hidden lg:table-cell">Joined</TableHead>
                  <TableHead className="text-right pr-4">Actions</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {loading ? (
                  Array.from({ length: 10 }).map((_, i) => (
                    <TableRow key={i}>
                      <TableCell className="pl-4"><Skeleton className="h-4 w-4" /></TableCell>
                      <TableCell><Skeleton className="h-8 w-40" /></TableCell>
                      <TableCell className="hidden md:table-cell"><Skeleton className="h-4 w-36" /></TableCell>
                      <TableCell><Skeleton className="h-5 w-16" /></TableCell>
                      <TableCell><Skeleton className="h-5 w-16" /></TableCell>
                      <TableCell className="hidden lg:table-cell"><Skeleton className="h-4 w-20" /></TableCell>
                      <TableCell className="text-right pr-4"><Skeleton className="h-8 w-8 ml-auto" /></TableCell>
                    </TableRow>
                  ))
                ) : users.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={7} className="py-12">
                      <EmptyState
                        icon={Users}
                        title="No users found"
                        description={search || roleFilter !== 'all' || statusFilter !== 'all'
                          ? 'Try adjusting your filters.'
                          : 'No users have registered yet.'}
                      />
                    </TableCell>
                  </TableRow>
                ) : (
                  users.map((u) => {
                    const displayName = u.clientProfile?.displayName || u.freelancerProfile?.displayName || u.name || 'Unknown'
                    return (
                      <TableRow key={u.id} className="group">
                        <TableCell className="pl-4">
                          <Checkbox
                            checked={selectedUsers.has(u.id)}
                            onCheckedChange={() => toggleSelect(u.id)}
                          />
                        </TableCell>
                        <TableCell>
                          <div className="flex items-center gap-3">
                            <Avatar className="h-8 w-8">
                              <AvatarFallback className={cn(
                                'text-xs',
                                u.role === 'FREELANCER' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'
                              )}>
                                {displayName.slice(0, 2).toUpperCase()}
                              </AvatarFallback>
                            </Avatar>
                            <span className="text-sm font-medium text-slate-900">{displayName}</span>
                          </div>
                        </TableCell>
                        <TableCell className="hidden md:table-cell text-sm text-slate-500">{u.email}</TableCell>
                        <TableCell>
                          <Badge
                            variant="secondary"
                            className={cn(
                              'text-xs font-medium',
                              u.role === 'CLIENT' ? 'bg-amber-100 text-amber-700'
                                : u.role === 'ADMIN' ? 'bg-red-100 text-red-700'
                                : 'bg-emerald-100 text-emerald-700'
                            )}
                          >
                            {u.role}
                          </Badge>
                        </TableCell>
                        <TableCell>
                          <div className="flex items-center gap-1.5">
                            {u.isVerified && (
                              <Badge variant="secondary" className="bg-emerald-100 text-emerald-700 text-xs">
                                <CheckCircle2 className="w-3 h-3 mr-0.5" /> Verified
                              </Badge>
                            )}
                            {u.isSuspended && (
                              <Badge variant="secondary" className="bg-red-100 text-red-700 text-xs">
                                <Ban className="w-3 h-3 mr-0.5" /> Suspended
                              </Badge>
                            )}
                            {!u.isVerified && !u.isSuspended && (
                              <span className="text-xs text-slate-400">Active</span>
                            )}
                          </div>
                        </TableCell>
                        <TableCell className="hidden lg:table-cell text-sm text-slate-500">
                          {new Date(u.createdAt).toLocaleDateString()}
                        </TableCell>
                        <TableCell className="text-right pr-4">
                          <DropdownMenu>
                            <DropdownMenuTrigger asChild>
                              <Button variant="ghost" size="icon" className="h-8 w-8">
                                <MoreVertical className="w-4 h-4" />
                              </Button>
                            </DropdownMenuTrigger>
                            <DropdownMenuContent align="end">
                              <DropdownMenuItem onClick={() => navigate('freelancers/detail', { id: u.id })}>
                                <Eye className="w-3.5 h-3.5 mr-2" /> View Profile
                              </DropdownMenuItem>
                              <DropdownMenuSeparator />
                              {!u.isVerified && (
                                <DropdownMenuItem onClick={() => handleAction(u.id, 'verify')}>
                                  <Shield className="w-3.5 h-3.5 mr-2" /> Verify User
                                </DropdownMenuItem>
                              )}
                              {u.isVerified && (
                                <DropdownMenuItem onClick={() => handleAction(u.id, 'verify')} className="text-slate-500">
                                  <ShieldOff className="w-3.5 h-3.5 mr-2" /> Unverify
                                </DropdownMenuItem>
                              )}
                              {!u.isSuspended ? (
                                <DropdownMenuItem
                                  onClick={() => {
                                    setSuspendTarget(u.id)
                                    setSuspendDialogOpen(true)
                                  }}
                                  className="text-red-600"
                                >
                                  <Ban className="w-3.5 h-3.5 mr-2" /> Suspend User
                                </DropdownMenuItem>
                              ) : (
                                <DropdownMenuItem onClick={() => handleAction(u.id, 'unsuspend')}>
                                  <CheckCircle2 className="w-3.5 h-3.5 mr-2" /> Unsuspend User
                                </DropdownMenuItem>
                              )}
                            </DropdownMenuContent>
                          </DropdownMenu>
                        </TableCell>
                      </TableRow>
                    )
                  })
                )}
              </TableBody>
            </Table>
          </div>

          {/* Pagination */}
          {totalPages > 1 && (
            <div className="flex items-center justify-between px-4 py-3 border-t border-slate-100">
              <span className="text-sm text-slate-500">Page {page} of {totalPages}</span>
              <div className="flex items-center gap-1">
                <Button
                  variant="outline"
                  size="icon"
                  className="h-8 w-8"
                  onClick={() => setPage(p => Math.max(1, p - 1))}
                  disabled={page <= 1}
                >
                  <ChevronLeft className="w-4 h-4" />
                </Button>
                <Button
                  variant="outline"
                  size="icon"
                  className="h-8 w-8"
                  onClick={() => setPage(p => Math.min(totalPages, p + 1))}
                  disabled={page >= totalPages}
                >
                  <ChevronRight className="w-4 h-4" />
                </Button>
              </div>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Suspend Dialog */}
      <Dialog open={suspendDialogOpen} onOpenChange={setSuspendDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle className="text-red-600">Suspend User</DialogTitle>
          </DialogHeader>
          <div className="space-y-4 mt-4">
            <Textarea
              placeholder="Reason for suspension (optional but recommended)..."
              value={suspendReason}
              onChange={e => setSuspendReason(e.target.value)}
              rows={3}
            />
            <div className="flex justify-end gap-2">
              <Button variant="outline" onClick={() => setSuspendDialogOpen(false)}>Cancel</Button>
              <Button
                variant="destructive"
                onClick={() => suspendTarget && handleAction(suspendTarget, 'suspend', suspendReason)}
                disabled={actionLoading}
              >
                Suspend User
              </Button>
            </div>
          </div>
        </DialogContent>
      </Dialog>
    </motion.div>
  )
}
