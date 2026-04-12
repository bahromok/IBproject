'use client'

import { useEffect, useState } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { Shield, X } from 'lucide-react'
import { cn } from '@/lib/utils'

const tips = [
  { icon: '🔒', text: 'Your identity is protected. Other users can only see your display name.' },
  { icon: '💰', text: 'Never share payment info outside the platform.' },
  { icon: '⚠️', text: 'If something seems suspicious, report it immediately.' },
]

export function SafetyBanner({ className }: { className?: string }) {
  const [dismissed, setDismissed] = useState(() => {
    if (typeof window === 'undefined') return false
    return localStorage.getItem('safety-banner-dismissed') === 'true'
  })
  const [tipIndex, setTipIndex] = useState(0)

  useEffect(() => {
    if (dismissed) return
    const interval = setInterval(() => {
      setTipIndex((prev) => (prev + 1) % tips.length)
    }, 6000)
    return () => clearInterval(interval)
  }, [dismissed])

  const handleDismiss = () => {
    setDismissed(true)
    localStorage.setItem('safety-banner-dismissed', 'true')
  }

  if (dismissed) return null

  const tip = tips[tipIndex]

  return (
    <AnimatePresence mode="wait">
      <motion.div
        key={tipIndex}
        initial={{ opacity: 0, y: -5 }}
        animate={{ opacity: 1, y: 0 }}
        exit={{ opacity: 0, y: -5 }}
        transition={{ duration: 0.3 }}
        className={cn(
          'relative flex items-center gap-3 px-4 py-2.5 rounded-lg bg-emerald-50 border border-emerald-200',
          className
        )}
      >
        <Shield className="w-4 h-4 text-emerald-600 shrink-0" />
        <p className="flex-1 text-xs text-emerald-800">
          <span className="mr-1.5">{tip.icon}</span>
          {tip.text}
        </p>
        <button
          onClick={handleDismiss}
          className="shrink-0 p-0.5 rounded hover:bg-emerald-100 transition-colors"
          aria-label="Dismiss safety tip"
        >
          <X className="w-3.5 h-3.5 text-emerald-500" />
        </button>
      </motion.div>
    </AnimatePresence>
  )
}
