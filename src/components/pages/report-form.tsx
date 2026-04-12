'use client'

import { useState } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Label } from '@/components/ui/label'
import { RadioGroup, RadioGroupItem } from '@/components/ui/radio-group'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select'
import {
  ShieldCheck,
  Upload,
  AlertTriangle,
  CheckCircle2,
  Flag,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'

const REPORT_TYPES = [
  { value: 'SCAM', label: 'Scam or Fraudulent Activity', emoji: '🚫', color: 'text-red-600' },
  { value: 'HARASSMENT', label: 'Harassment or Intimidation', emoji: '⚠️', color: 'text-amber-600' },
  { value: 'SPAM', label: 'Spam or Unsolicited Contact', emoji: '📧', color: 'text-orange-600' },
  { value: 'INAPPROPRIATE_CONTENT', label: 'Inappropriate Content', emoji: '🚷', color: 'text-purple-600' },
  { value: 'PAYMENT_DISPUTE', label: 'Payment Dispute', emoji: '💰', color: 'text-emerald-600' },
  { value: 'OTHER', label: 'Other', emoji: '📋', color: 'text-slate-600' },
]

export function ReportForm() {
  const { pageParams } = useNavigationStore()
  const [reportType, setReportType] = useState('')
  const [description, setDescription] = useState('')
  const [targetId, setTargetId] = useState(pageParams?.userId || pageParams?.targetId || '')
  const [targetType, setTargetType] = useState('user')
  const [submitting, setSubmitting] = useState(false)
  const [submitted, setSubmitted] = useState(false)
  const [error, setError] = useState('')

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError('')

    if (!reportType) {
      setError('Please select a report type.')
      return
    }
    if (description.length < 50) {
      setError('Description must be at least 50 characters.')
      return
    }
    if (!targetId) {
      setError('Please specify who or what you are reporting.')
      return
    }

    setSubmitting(true)
    try {
      const res = await authFetch('/api/reports', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          targetType,
          targetId,
          reportType,
          description,
        }),
      })

      if (res.ok) {
        setSubmitted(true)
      } else {
        const data = await res.json()
        setError(data.error || 'Failed to submit report. Please try again.')
      }
    } catch {
      setError('An error occurred. Please try again.')
    } finally {
      setSubmitting(false)
    }
  }

  const { navigate } = useNavigationStore()

  if (submitted) {
    return (
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="max-w-lg mx-auto"
      >
        <Card className="border-emerald-200 bg-emerald-50/50">
          <CardContent className="pt-8 pb-8 text-center">
            <div className="w-16 h-16 rounded-full bg-emerald-100 flex items-center justify-center mx-auto mb-4">
              <CheckCircle2 className="w-8 h-8 text-emerald-600" />
            </div>
            <h2 className="text-xl font-bold text-slate-900 mb-2">Report Submitted Successfully</h2>
            <p className="text-slate-600 mb-6">
              Your report has been submitted. Our team will review it within 24-48 hours.
              You&apos;ll receive a notification once a decision has been made.
            </p>
            <div className="bg-white rounded-lg p-4 border border-emerald-100 mb-6">
              <div className="flex items-center gap-2 text-sm text-slate-600">
                <ShieldCheck className="w-4 h-4 text-emerald-500" />
                <span>Your identity will remain confidential throughout this process.</span>
              </div>
            </div>
            <Button
              variant="outline"
              onClick={() => {
                setSubmitted(false)
                setReportType('')
                setDescription('')
                setTargetId('')
                navigate('dashboard')
              }}
            >
              Go to Dashboard
            </Button>
          </CardContent>
        </Card>
      </motion.div>
    )
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="max-w-2xl mx-auto"
    >
      <div className="mb-6">
        <div className="flex items-center gap-3 mb-2">
          <div className="w-10 h-10 rounded-xl bg-red-100 flex items-center justify-center">
            <Flag className="w-5 h-5 text-red-600" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-900">File a Report</h1>
            <p className="text-sm text-slate-500">Help us keep StudentHire safe for everyone</p>
          </div>
        </div>
      </div>

      <Card>
        <CardContent className="pt-6">
          <form onSubmit={handleSubmit} className="space-y-6">
            {/* Report Type */}
            <div className="space-y-3">
              <Label className="text-sm font-semibold text-slate-900">Report Type</Label>
              <RadioGroup value={reportType} onValueChange={setReportType} className="space-y-2">
                {REPORT_TYPES.map((type) => (
                  <label
                    key={type.value}
                    htmlFor={type.value}
                    className={cn(
                      'flex items-center gap-3 p-3 rounded-lg border-2 cursor-pointer transition-all',
                      reportType === type.value
                        ? 'border-emerald-500 bg-emerald-50/50'
                        : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50'
                    )}
                  >
                    <RadioGroupItem value={type.value} id={type.value} />
                    <span className="text-lg">{type.emoji}</span>
                    <span className={cn('text-sm font-medium', reportType === type.value ? 'text-emerald-900' : 'text-slate-700')}>
                      {type.label}
                    </span>
                  </label>
                ))}
              </RadioGroup>
            </div>

            {/* Target Type & ID */}
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div className="space-y-2">
                <Label htmlFor="targetType" className="text-sm font-semibold text-slate-900">
                  What are you reporting?
                </Label>
                <Select value={targetType} onValueChange={setTargetType}>
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="user">A User</SelectItem>
                    <SelectItem value="job">A Job Posting</SelectItem>
                    <SelectItem value="proposal">A Proposal</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-2">
                <Label htmlFor="targetId" className="text-sm font-semibold text-slate-900">
                  {targetType === 'user' ? 'User ID' : targetType === 'job' ? 'Job ID' : 'Proposal ID'}
                </Label>
                <Input
                  id="targetId"
                  placeholder={targetType === 'user' ? 'Enter user ID...' : targetType === 'job' ? 'Enter job ID...' : 'Enter proposal ID...'}
                  value={targetId}
                  onChange={e => setTargetId(e.target.value)}
                />
              </div>
            </div>

            {/* Description */}
            <div className="space-y-2">
              <Label htmlFor="description" className="text-sm font-semibold text-slate-900">
                Description <span className="text-red-500">*</span>
              </Label>
              <Textarea
                id="description"
                placeholder="Please provide a detailed description of the issue. Include relevant dates, actions, and any other details that would help our team investigate..."
                rows={5}
                value={description}
                onChange={e => setDescription(e.target.value)}
                className="resize-none"
              />
              <div className="flex justify-between items-center">
                <p className="text-xs text-slate-400">
                  Minimum 50 characters
                </p>
                <span className={cn(
                  'text-xs font-medium',
                  description.length >= 50 ? 'text-emerald-600' : 'text-slate-400'
                )}>
                  {description.length} / 5000
                </span>
              </div>
            </div>

            {/* Evidence Upload Placeholder */}
            <div className="space-y-2">
              <Label className="text-sm font-semibold text-slate-900">Evidence (Optional)</Label>
              <div className="border-2 border-dashed border-slate-200 rounded-lg p-6 text-center hover:border-emerald-300 transition-colors cursor-pointer">
                <Upload className="w-8 h-8 text-slate-400 mx-auto mb-2" />
                <p className="text-sm text-slate-600 font-medium">Click to upload files</p>
                <p className="text-xs text-slate-400 mt-1">Screenshots, documents, or other evidence</p>
              </div>
            </div>

            {/* Error */}
            {error && (
              <motion.div
                initial={{ opacity: 0, y: -5 }}
                animate={{ opacity: 1, y: 0 }}
                className="flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg"
              >
                <AlertTriangle className="w-4 h-4 text-red-500 shrink-0" />
                <p className="text-sm text-red-700">{error}</p>
              </motion.div>
            )}

            {/* Submit */}
            <div className="flex items-center gap-3 pt-2">
              <Button
                type="submit"
                disabled={submitting || !reportType || description.length < 50 || !targetId}
                className="bg-red-600 hover:bg-red-700"
              >
                {submitting ? 'Submitting...' : 'Submit Report'}
              </Button>
              <p className="text-xs text-slate-400">
                All reports are reviewed confidentially by our trust & safety team.
              </p>
            </div>
          </form>
        </CardContent>
      </Card>
    </motion.div>
  )
}
