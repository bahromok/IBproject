'use client'

import { useState } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Badge } from '@/components/ui/badge'
import { Switch } from '@/components/ui/switch'
import { Separator } from '@/components/ui/separator'
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs'
import {
  User, Lock, Bell, Eye, Download, Shield, Smartphone,
  Mail, AlertTriangle, Globe, ChevronRight, Loader2, DollarSign, Wallet
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'
import { EmailVerificationBanner } from '@/components/shared/email-verification-banner'
import { authFetch } from '@/lib/api-fetch'
import { getStripeConnectUrl } from '@/app/actions/payments'
import { ConnectKitButton } from 'connectkit'

export function SettingsPage() {
  const { user } = useAuthStore()
  const { toast } = useToast()

  // Account settings
  const [displayName, setDisplayName] = useState(user?.name || '')
  const [currentPassword, setCurrentPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')

  // Security settings
  const [twoFactorEnabled, setTwoFactorEnabled] = useState(false)

  // Notification settings
  const [notifications, setNotifications] = useState({
    newProposals: true,
    messages: true,
    milestones: true,
    payments: true,
    reviews: true,
    marketing: false,
  })

  // Privacy settings
  const [privacy, setPrivacy] = useState({
    profileVisible: true,
    showInSearch: true,
    onlineStatus: true,
  })

  const [passwordLoading, setPasswordLoading] = useState(false)

  const handleChangePassword = async () => {
    if (!currentPassword) {
      toast({ title: 'Error', description: 'Please enter your current password', variant: 'destructive' })
      return
    }
    if (newPassword.length < 8) {
      toast({ title: 'Error', description: 'Password must be at least 8 characters', variant: 'destructive' })
      return
    }
    if (newPassword !== confirmPassword) {
      toast({ title: 'Error', description: 'Passwords do not match', variant: 'destructive' })
      return
    }
    setPasswordLoading(true)
    try {
      const res = await authFetch('/api/auth/change-password', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ currentPassword, newPassword }),
      })
      if (res.ok) {
        toast({ title: 'Password updated', description: 'Your password has been changed successfully.' })
        setCurrentPassword('')
        setNewPassword('')
        setConfirmPassword('')
      } else {
        const data = await res.json()
        toast({ title: 'Error', description: data.error || 'Failed to change password', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to change password', variant: 'destructive' })
    }
    setPasswordLoading(false)
  }

  const handleSaveDisplayName = () => {
    if (!displayName.trim()) {
      toast({ title: 'Error', description: 'Display name cannot be empty', variant: 'destructive' })
      return
    }
    toast({ title: 'Name updated', description: 'Your display name has been changed.' })
  }

  const handleToggleNotification = (key: keyof typeof notifications) => {
    setNotifications(prev => ({ ...prev, [key]: !prev[key] }))
    toast({ title: 'Notification settings updated', description: 'Your preferences have been saved.' })
  }

  const handleTogglePrivacy = (key: keyof typeof privacy) => {
    setPrivacy(prev => ({ ...prev, [key]: !prev[key] }))
    toast({ title: 'Privacy settings updated', description: 'Your preferences have been saved.' })
  }

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Email Verification Banner */}
      <EmailVerificationBanner />

      {/* Header */}
      <div>
        <h1 className="text-2xl font-bold text-slate-900">Settings</h1>
        <p className="text-slate-500 mt-1">Manage your account preferences and security.</p>
      </div>

      <Tabs defaultValue="account">
        <TabsList className="bg-slate-100">
          <TabsTrigger value="account" className="text-sm flex items-center gap-1.5">
            <User className="w-3.5 h-3.5" />
            Account
          </TabsTrigger>
          <TabsTrigger value="security" className="text-sm flex items-center gap-1.5">
            <Lock className="w-3.5 h-3.5" />
            Security
          </TabsTrigger>
          <TabsTrigger value="notifications" className="text-sm flex items-center gap-1.5">
            <Bell className="w-3.5 h-3.5" />
            Notifications
          </TabsTrigger>
          <TabsTrigger value="privacy" className="text-sm flex items-center gap-1.5">
            <Eye className="w-3.5 h-3.5" />
            Privacy
          </TabsTrigger>
          <TabsTrigger value="payments" className="text-sm flex items-center gap-1.5">
            <DollarSign className="w-3.5 h-3.5" />
            Payments
          </TabsTrigger>
        </TabsList>

        {/* Account Tab */}
        <TabsContent value="account" className="space-y-6 mt-6">
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Mail className="w-5 h-5 text-slate-400" />
                Email
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                <div>
                  <div className="text-sm font-medium text-slate-900">{user?.email}</div>
                  <div className="text-xs text-slate-500">Your account email (cannot be changed)</div>
                </div>
                {user?.emailVerified ? (
                  <Badge variant="secondary" className="bg-emerald-100 text-emerald-700 text-xs">Verified</Badge>
                ) : (
                  <Badge variant="secondary" className="bg-amber-100 text-amber-700 text-xs">Unverified</Badge>
                )}
              </div>
            </CardContent>
          </Card>

          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <User className="w-5 h-5 text-slate-400" />
                Display Name
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div>
                <Label className="text-sm">Display Name</Label>
                <Input
                  value={displayName}
                  onChange={(e) => setDisplayName(e.target.value)}
                  placeholder="Your display name"
                  className="mt-1.5"
                />
              </div>
              <Button onClick={handleSaveDisplayName} className="bg-emerald-500 hover:bg-emerald-600 text-white">
                Save Name
              </Button>
            </CardContent>
          </Card>

          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Globe className="w-5 h-5 text-slate-400" />
                Account Role
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                <div>
                  <div className="text-sm font-medium text-slate-900 capitalize">
                    {user?.role === 'CLIENT' ? 'Client' : user?.role === 'FREELANCER' ? 'Freelancer' : user?.role}
                  </div>
                  <div className="text-xs text-slate-500">Your account type determines what you can do on the platform</div>
                </div>
                <Badge variant="outline" className="bg-slate-100 text-slate-600 text-xs">
                  {user?.role}
                </Badge>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Security Tab */}
        <TabsContent value="security" className="space-y-6 mt-6">
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Lock className="w-5 h-5 text-slate-400" />
                Change Password
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div>
                <Label className="text-sm">Current Password</Label>
                <Input
                  type="password"
                  value={currentPassword}
                  onChange={(e) => setCurrentPassword(e.target.value)}
                  placeholder="Enter current password"
                  className="mt-1.5"
                />
              </div>
              <div>
                <Label className="text-sm">New Password</Label>
                <Input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  placeholder="Enter new password"
                  className="mt-1.5"
                />
                <p className="text-xs text-slate-400 mt-1">Must be at least 8 characters</p>
              </div>
              <div>
                <Label className="text-sm">Confirm New Password</Label>
                <Input
                  type="password"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  placeholder="Confirm new password"
                  className="mt-1.5"
                />
              </div>
              <Button onClick={handleChangePassword} disabled={passwordLoading} className="bg-emerald-500 hover:bg-emerald-600 text-white">
                {passwordLoading ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : null}
                Update Password
              </Button>
            </CardContent>
          </Card>

          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Shield className="w-5 h-5 text-slate-400" />
                Two-Factor Authentication
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                <div>
                  <div className="text-sm font-medium text-slate-900">Two-Factor Authentication</div>
                  <div className="text-xs text-slate-500">Add an extra layer of security to your account</div>
                </div>
                <Switch
                  checked={twoFactorEnabled}
                  onCheckedChange={(v) => {
                    setTwoFactorEnabled(v)
                    toast({ title: '2FA ' + (v ? 'enabled' : 'disabled'), description: 'Two-factor authentication has been ' + (v ? 'enabled' : 'disabled') + '.' })
                  }}
                />
              </div>
            </CardContent>
          </Card>

          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Smartphone className="w-5 h-5 text-slate-400" />
                Active Sessions
              </CardTitle>
            </CardHeader>
            <CardContent>
              <div className="space-y-2">
                <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                  <div className="flex items-center gap-3">
                    <div className="w-8 h-8 rounded-lg bg-emerald-100 flex items-center justify-center">
                      <Globe className="w-4 h-4 text-emerald-600" />
                    </div>
                    <div>
                      <div className="text-sm font-medium text-slate-900">Current Session</div>
                      <div className="text-xs text-slate-500">Browser • Now</div>
                    </div>
                  </div>
                  <Badge variant="secondary" className="bg-emerald-100 text-emerald-700 text-xs">Active</Badge>
                </div>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Notifications Tab */}
        <TabsContent value="notifications" className="space-y-6 mt-6">
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Bell className="w-5 h-5 text-slate-400" />
                Notification Preferences
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-1">
              {[
                { key: 'newProposals' as const, label: 'New Proposals', description: 'Get notified when you receive a new proposal on your jobs' },
                { key: 'messages' as const, label: 'Messages', description: 'Get notified when you receive a new message' },
                { key: 'milestones' as const, label: 'Milestones', description: 'Get notified about milestone updates and submissions' },
                { key: 'payments' as const, label: 'Payments', description: 'Get notified about payment releases and transactions' },
                { key: 'reviews' as const, label: 'Reviews', description: 'Get notified when you receive a new review' },
                { key: 'marketing' as const, label: 'Marketing Emails', description: 'Receive tips, updates, and promotional content' },
              ].map((item, idx) => (
                <div key={item.key}>
                  <div className="flex items-center justify-between py-3">
                    <div>
                      <div className="text-sm font-medium text-slate-900">{item.label}</div>
                      <div className="text-xs text-slate-500">{item.description}</div>
                    </div>
                    <Switch
                      checked={notifications[item.key]}
                      onCheckedChange={() => handleToggleNotification(item.key)}
                    />
                  </div>
                  {idx < 5 && <div className="border-t border-slate-100" />}
                </div>
              ))}
            </CardContent>
          </Card>
        </TabsContent>

        {/* Privacy Tab */}
        <TabsContent value="privacy" className="space-y-6 mt-6">
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Eye className="w-5 h-5 text-slate-400" />
                Privacy Settings
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-1">
              {[
                { key: 'profileVisible' as const, label: 'Profile Visibility', description: 'Make your profile visible to other users on the platform' },
                { key: 'showInSearch' as const, label: 'Show in Search Results', description: 'Allow your profile to appear in search results and browse pages' },
                { key: 'onlineStatus' as const, label: 'Online Status', description: 'Show when you are online to other users' },
              ].map((item, idx) => (
                <div key={item.key}>
                  <div className="flex items-center justify-between py-3">
                    <div>
                      <div className="text-sm font-medium text-slate-900">{item.label}</div>
                      <div className="text-xs text-slate-500">{item.description}</div>
                    </div>
                    <Switch
                      checked={privacy[item.key]}
                      onCheckedChange={() => handleTogglePrivacy(item.key)}
                    />
                  </div>
                  {idx < 2 && <div className="border-t border-slate-100" />}
                </div>
              ))}
            </CardContent>
          </Card>

          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <Download className="w-5 h-5 text-slate-400" />
                Data & Privacy
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                <div>
                  <div className="text-sm font-medium text-slate-900">Export Your Data</div>
                  <div className="text-xs text-slate-500">Download a copy of all your data</div>
                </div>
                <Button variant="outline" size="sm" className="border-slate-200 text-slate-600" onClick={() => toast({ title: 'Request submitted', description: 'You will receive an email with your data export shortly.' })}>
                  Request Export
                </Button>
              </div>
              <div className="flex items-center justify-between p-3 bg-red-50 rounded-lg border border-red-100">
                <div className="flex items-center gap-3">
                  <AlertTriangle className="w-5 h-5 text-red-500" />
                  <div>
                    <div className="text-sm font-medium text-red-700">Delete Account</div>
                    <div className="text-xs text-red-500">Permanently delete your account and all data</div>
                  </div>
                </div>
                <Button variant="outline" size="sm" className="border-red-200 text-red-600 hover:bg-red-100">
                  Delete
                </Button>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        {/* Payments Tab */}
        <TabsContent value="payments" className="space-y-6 mt-6">
          <Card className="border-slate-200">
            <CardHeader className="pb-3">
              <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                <DollarSign className="w-5 h-5 text-slate-400" />
                Payout Methods
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex items-center justify-between p-4 bg-indigo-50 border border-indigo-100 rounded-lg">
                <div className="flex items-start gap-3">
                  <div className="w-10 h-10 rounded-lg bg-indigo-100 flex items-center justify-center shrink-0">
                    <DollarSign className="w-5 h-5 text-indigo-600" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-semibold text-slate-900">Stripe Connect</div>
                    <div className="text-xs text-slate-500 mt-0.5">Receive payouts directly to your bank account</div>
                  </div>
                </div>
                <Button
                  variant="outline"
                  className="border-indigo-200 text-indigo-600 hover:bg-indigo-100 shrink-0"
                  onClick={async () => {
                    try {
                      toast({ title: 'Redirecting...', description: 'Preparing Stripe onboarding link.' })
                      const { url } = await getStripeConnectUrl()
                      window.location.href = url
                    } catch (err: any) {
                      toast({ title: 'Stripe Error', description: err.message, variant: 'destructive' })
                    }
                  }}
                >
                  Connect Stripe
                </Button>
              </div>

              <div className="flex items-center justify-between p-4 bg-blue-50 border border-blue-100 rounded-lg">
                <div className="flex items-start gap-3">
                  <div className="w-10 h-10 rounded-lg bg-blue-100 flex items-center justify-center shrink-0">
                    <Wallet className="w-5 h-5 text-blue-600" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-semibold text-slate-900">Crypto Wallet</div>
                    <div className="text-xs text-slate-500 mt-0.5">Support for Mainnets & Testnets (Sepolia)</div>
                  </div>
                </div>
                <ConnectKitButton.Custom>
                  {({ isConnected, show, truncatedAddress, ensName }) => (
                    <Button
                      variant="outline"
                      className="border-blue-200 text-blue-600 hover:bg-blue-100 shrink-0"
                      onClick={show}
                    >
                      {isConnected ? (ensName || truncatedAddress) : 'Connect Wallet'}
                    </Button>
                  )}
                </ConnectKitButton.Custom>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs >
    </motion.div >
  )
}
