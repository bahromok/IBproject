'use client'
import { useCountUp } from '@/hooks/use-count-up'
import { formatCurrency, formatNumber } from '@/lib/format'
import { LucideIcon } from 'lucide-react'
import { cn } from '@/lib/utils'
import { useEffect, useState } from 'react'
import { motion } from 'framer-motion'

interface StatCardProps {
  label: string
  value: number
  prefix?: string
  suffix?: string
  icon: LucideIcon
  format?: 'number' | 'currency'
  className?: string
  iconColor?: string
  sub?: string
  decimals?: number
}

export function StatCard({ label, value, prefix = '', suffix = '', icon: Icon, format = 'number', className, iconColor, sub, decimals = 0 }: StatCardProps) {
  const [shouldAnimate, setShouldAnimate] = useState(false)

  useEffect(() => {
    const timer = setTimeout(() => setShouldAnimate(true), 100)
    return () => clearTimeout(timer)
  }, [])

  const countValue = useCountUp(shouldAnimate ? value : 0, 1800, decimals)

  let displayValue: string
  if (format === 'currency') {
    displayValue = formatCurrency(countValue)
  } else {
    displayValue = `${prefix}${formatNumber(countValue)}${suffix}`
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      className={cn('bg-white rounded-xl border border-slate-200 p-5 hover:shadow-sm transition-shadow', className)}
    >
      <div className="flex items-center justify-between mb-3">
        <div className={cn('w-10 h-10 rounded-lg flex items-center justify-center', iconColor || 'bg-emerald-100 text-emerald-600')}>
          <Icon className="w-5 h-5" />
        </div>
      </div>
      <div className="text-2xl font-bold text-slate-900">{displayValue}</div>
      <div className="text-sm text-slate-500 mt-0.5">{label}</div>
      {sub && <div className="text-xs text-slate-400 mt-1">{sub}</div>}
    </motion.div>
  )
}
