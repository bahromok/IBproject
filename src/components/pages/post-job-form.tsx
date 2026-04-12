'use client'

import { useState } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Input } from '@/components/ui/input'
import { Textarea } from '@/components/ui/textarea'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Switch } from '@/components/ui/switch'
import { Separator } from '@/components/ui/separator'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'
import {
  ChevronLeft, ChevronRight, Check, Briefcase, DollarSign,
  Clock, X, Plus, AlertCircle
} from 'lucide-react'

const CATEGORIES = [
  { value: 'WEB_DEVELOPER', label: 'Web Dev' },
  { value: 'MOBILE_DEVELOPER', label: 'Mobile App' },
  { value: 'DESIGNER', label: 'Design' },
  { value: 'VIDEO_EDITOR', label: 'Video Edit' },
  { value: 'CONTENT_CREATOR', label: 'Content' },
  { value: 'MARKETING', label: 'Marketing' },
  { value: 'WRITING', label: 'Writing' },
  { value: 'TRANSLATION', label: 'Translation' },
  { value: 'DATA_SCIENCE', label: 'Data Science' },
  { value: 'ADMIN_SUPPORT', label: 'Admin' },
  { value: 'OTHER', label: 'Other' },
]

const EXPERIENCE_LEVELS = [
  { value: 'BEGINNER', label: 'Beginner' },
  { value: 'INTERMEDIATE', label: 'Intermediate' },
  { value: 'ADVANCED', label: 'Advanced' },
  { value: 'EXPERT', label: 'Expert' },
]

const DURATIONS = [
  { value: 'LESS_THAN_1_WEEK', label: 'Less than 1 week' },
  { value: 'ONE_TO_4_WEEKS', label: '1-4 weeks' },
  { value: 'ONE_TO_3_MONTHS', label: '1-3 months' },
  { value: 'THREE_TO_6_MONTHS', label: '3-6 months' },
  { value: 'MORE_THAN_6_MONTHS', label: '6+ months' },
]

const COMMON_SKILLS = [
  'React', 'JavaScript', 'TypeScript', 'Python', 'Node.js', 'Next.js',
  'CSS', 'HTML', 'Figma', 'Photoshop', 'Illustrator', 'Premiere Pro',
  'After Effects', 'WordPress', 'Shopify', 'SEO', 'Copywriting',
  'Graphic Design', 'UI/UX', 'Mobile Development', 'Flutter',
  'AWS', 'Docker', 'MongoDB', 'PostgreSQL',
]

export default function PostJobForm() {
  const { goBack, navigate } = useNavigationStore()
  const { toast } = useToast()

  const [step, setStep] = useState(1)
  const [submitting, setSubmitting] = useState(false)

  // Form state
  const [title, setTitle] = useState('')
  const [category, setCategory] = useState('')
  const [description, setDescription] = useState('')
  const [skills, setSkills] = useState<string[]>([])
  const [skillInput, setSkillInput] = useState('')

  const [budgetType, setBudgetType] = useState<'FIXED_PRICE' | 'HOURLY'>('FIXED_PRICE')
  const [budgetMin, setBudgetMin] = useState('')
  const [budgetMax, setBudgetMax] = useState('')
  const [duration, setDuration] = useState('')
  const [experienceLevel, setExperienceLevel] = useState('')
  const [isUrgent, setIsUrgent] = useState(false)

  const [errors, setErrors] = useState<Record<string, string>>({})

  const addSkill = (skill: string) => {
    const trimmed = skill.trim()
    if (trimmed && !skills.includes(trimmed) && skills.length < 20) {
      setSkills([...skills, trimmed])
      setSkillInput('')
    }
  }

  const removeSkill = (skill: string) => {
    setSkills(skills.filter(s => s !== skill))
  }

  const validateStep = (s: number): boolean => {
    const newErrors: Record<string, string> = {}

    if (s === 1) {
      if (!title.trim()) newErrors.title = 'Title is required'
      if (!category) newErrors.category = 'Category is required'
      if (!description.trim()) newErrors.description = 'Description is required'
      else if (description.trim().length < 50) newErrors.description = 'Description must be at least 50 characters'
    }

    if (s === 2) {
      if (!budgetMin) newErrors.budgetMin = 'Budget minimum is required'
      else if (parseFloat(budgetMin) <= 0) newErrors.budgetMin = 'Budget must be positive'
      if (budgetMax && parseFloat(budgetMax) < parseFloat(budgetMin))
        newErrors.budgetMax = 'Max must be greater than or equal to min'
      if (!duration) newErrors.duration = 'Duration is required'
      if (!experienceLevel) newErrors.experienceLevel = 'Experience level is required'
    }

    setErrors(newErrors)
    return Object.keys(newErrors).length === 0
  }

  const nextStep = () => {
    if (validateStep(step)) {
      setStep(s => Math.min(3, s + 1))
    }
  }

  const prevStep = () => setStep(s => Math.max(1, s - 1))

  const submitJob = async (asDraft: boolean) => {
    setSubmitting(true)
    try {
      const res = await authFetch('/api/jobs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: title.trim(),
          category,
          description: description.trim(),
          skills,
          budgetType,
          budgetMin: parseFloat(budgetMin),
          budgetMax: budgetMax ? parseFloat(budgetMax) : null,
          duration,
          experienceLevel,
          isUrgent,
          status: asDraft ? 'DRAFT' : 'open',
        }),
      })
      const data = await res.json()
      if (res.ok) {
        toast({
          title: asDraft ? 'Draft Saved!' : 'Job Published!',
          description: asDraft ? 'Your job has been saved as a draft.' : 'Your job is now live and accepting proposals.',
        })
        navigate('jobs')
      } else {
        toast({ title: 'Error', description: data.error || 'Failed to create job', variant: 'destructive' })
      }
    } catch {
      toast({ title: 'Error', description: 'Failed to create job. Please try again.', variant: 'destructive' })
    } finally {
      setSubmitting(false)
    }
  }

  const steps = [
    { num: 1, label: 'Job Details' },
    { num: 2, label: 'Budget & Timeline' },
    { num: 3, label: 'Review & Post' },
  ]

  return (
    <div className="min-h-screen bg-gray-50/50">
      {/* Top Bar */}
      <div className="sticky top-0 z-40 bg-white border-b border-border">
        <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-14">
            <div className="flex items-center gap-4">
              <button onClick={goBack} className="text-muted-foreground hover:text-foreground transition-colors">
                <ChevronLeft className="w-5 h-5" />
              </button>
              <div className="flex items-center gap-2">
                <Briefcase className="w-5 h-5 text-emerald-600" />
                <h1 className="text-lg font-semibold">Post a Job</h1>
              </div>
            </div>
            <div className="flex items-center gap-2">
              {steps.map(s => (
                <div
                  key={s.num}
                  className={`flex items-center gap-2 px-3 py-1 rounded-full text-xs font-medium transition-colors ${step === s.num
                      ? 'bg-emerald-50 text-emerald-700'
                      : step > s.num
                        ? 'bg-emerald-100 text-emerald-600'
                        : 'bg-gray-100 text-muted-foreground'
                    }`}
                >
                  <span className="w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold bg-current text-white">
                    {step > s.num ? <Check className="w-3 h-3" /> : s.num}
                  </span>
                  <span className="hidden sm:inline">{s.label}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <AnimatePresence mode="wait">
          {/* Step 1: Job Details */}
          {step === 1 && (
            <motion.div
              key="step1"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
            >
              <Card className="border-0 shadow-sm">
                <CardHeader>
                  <CardTitle className="text-xl">Job Details</CardTitle>
                </CardHeader>
                <CardContent className="space-y-6">
                  {/* Title */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      Job Title <span className="text-destructive">*</span>
                    </label>
                    <Input
                      placeholder="e.g., Build a responsive e-commerce website"
                      value={title}
                      onChange={(e) => setTitle(e.target.value)}
                      maxLength={100}
                      className={errors.title ? 'border-destructive' : ''}
                    />
                    {errors.title && (
                      <p className="text-xs text-destructive flex items-center gap-1">
                        <AlertCircle className="w-3 h-3" /> {errors.title}
                      </p>
                    )}
                    <p className="text-xs text-muted-foreground">{title.length}/100 characters</p>
                  </div>

                  {/* Category */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      Category <span className="text-destructive">*</span>
                    </label>
                    <Select value={category} onValueChange={setCategory}>
                      <SelectTrigger className={errors.category ? 'border-destructive' : ''}>
                        <SelectValue placeholder="Select a category" />
                      </SelectTrigger>
                      <SelectContent>
                        {CATEGORIES.map(c => (
                          <SelectItem key={c.value} value={c.value}>{c.label}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    {errors.category && (
                      <p className="text-xs text-destructive flex items-center gap-1">
                        <AlertCircle className="w-3 h-3" /> {errors.category}
                      </p>
                    )}
                  </div>

                  {/* Description */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      Description <span className="text-destructive">*</span>
                    </label>
                    <Textarea
                      placeholder="Describe your project in detail. What needs to be done? What are the requirements and expectations? The more detail you provide, the better proposals you'll receive."
                      value={description}
                      onChange={(e) => setDescription(e.target.value)}
                      rows={10}
                      maxLength={5000}
                      className={`resize-none ${errors.description ? 'border-destructive' : ''}`}
                    />
                    {errors.description && (
                      <p className="text-xs text-destructive flex items-center gap-1">
                        <AlertCircle className="w-3 h-3" /> {errors.description}
                      </p>
                    )}
                    <p className="text-xs text-muted-foreground">{description.length}/5000 characters</p>
                  </div>

                  {/* Skills */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">Skills Required</label>
                    <div className="flex gap-2">
                      <Input
                        placeholder="Type a skill and press Enter"
                        value={skillInput}
                        onChange={(e) => setSkillInput(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.preventDefault()
                            addSkill(skillInput)
                          }
                        }}
                      />
                      <Button
                        type="button"
                        variant="outline"
                        onClick={() => addSkill(skillInput)}
                      >
                        <Plus className="w-4 h-4" />
                      </Button>
                    </div>

                    {/* Common skills suggestions */}
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {COMMON_SKILLS.filter(s => !skills.includes(s)).slice(0, 12).map(skill => (
                        <button
                          key={skill}
                          onClick={() => addSkill(skill)}
                          className="text-xs px-2 py-1 rounded-full border text-muted-foreground hover:bg-emerald-50 hover:text-emerald-700 hover:border-emerald-200 transition-colors"
                        >
                          + {skill}
                        </button>
                      ))}
                    </div>

                    {skills.length > 0 && (
                      <div className="flex flex-wrap gap-1.5 mt-3">
                        {skills.map(skill => (
                          <Badge key={skill} variant="secondary" className="gap-1 pl-2.5">
                            {skill}
                            <button
                              onClick={() => removeSkill(skill)}
                              className="ml-0.5 hover:text-destructive transition-colors"
                            >
                              <X className="w-3 h-3" />
                            </button>
                          </Badge>
                        ))}
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>
            </motion.div>
          )}

          {/* Step 2: Budget & Timeline */}
          {step === 2 && (
            <motion.div
              key="step2"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
            >
              <Card className="border-0 shadow-sm">
                <CardHeader>
                  <CardTitle className="text-xl">Budget & Timeline</CardTitle>
                </CardHeader>
                <CardContent className="space-y-6">
                  {/* Budget Type */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">Budget Type</label>
                    <div className="grid grid-cols-2 gap-3">
                      <button
                        onClick={() => setBudgetType('FIXED_PRICE')}
                        className={`p-4 rounded-xl border-2 text-center transition-all ${budgetType === 'FIXED_PRICE'
                            ? 'border-emerald-500 bg-emerald-50'
                            : 'border-gray-200 hover:border-gray-300'
                          }`}
                      >
                        <DollarSign className={`w-5 h-5 mx-auto mb-2 ${budgetType === 'FIXED_PRICE' ? 'text-emerald-600' : 'text-muted-foreground'}`} />
                        <div className={`font-medium text-sm ${budgetType === 'FIXED_PRICE' ? 'text-emerald-700' : ''}`}>
                          Fixed Price
                        </div>
                        <div className="text-xs text-muted-foreground mt-1">
                          Set a total budget for the project
                        </div>
                      </button>
                      <button
                        onClick={() => setBudgetType('HOURLY')}
                        className={`p-4 rounded-xl border-2 text-center transition-all ${budgetType === 'HOURLY'
                            ? 'border-emerald-500 bg-emerald-50'
                            : 'border-gray-200 hover:border-gray-300'
                          }`}
                      >
                        <Clock className={`w-5 h-5 mx-auto mb-2 ${budgetType === 'HOURLY' ? 'text-emerald-600' : 'text-muted-foreground'}`} />
                        <div className={`font-medium text-sm ${budgetType === 'HOURLY' ? 'text-emerald-700' : ''}`}>
                          Hourly Rate
                        </div>
                        <div className="text-xs text-muted-foreground mt-1">
                          Pay by the hour
                        </div>
                      </button>
                    </div>
                  </div>

                  {/* Budget Range */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      {budgetType === 'FIXED_PRICE' ? 'Budget Range ($)' : 'Hourly Rate Range ($)'} <span className="text-destructive">*</span>
                    </label>
                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="text-xs text-muted-foreground">Min</label>
                        <Input
                          type="number"
                          placeholder="0"
                          value={budgetMin}
                          onChange={(e) => setBudgetMin(e.target.value)}
                          min="1"
                          className={errors.budgetMin ? 'border-destructive' : ''}
                        />
                        {errors.budgetMin && (
                          <p className="text-xs text-destructive mt-1">{errors.budgetMin}</p>
                        )}
                      </div>
                      <div>
                        <label className="text-xs text-muted-foreground">Max</label>
                        <Input
                          type="number"
                          placeholder="Optional"
                          value={budgetMax}
                          onChange={(e) => setBudgetMax(e.target.value)}
                          min="1"
                          className={errors.budgetMax ? 'border-destructive' : ''}
                        />
                        {errors.budgetMax && (
                          <p className="text-xs text-destructive mt-1">{errors.budgetMax}</p>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Duration */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      Project Duration <span className="text-destructive">*</span>
                    </label>
                    <Select value={duration} onValueChange={setDuration}>
                      <SelectTrigger className={errors.duration ? 'border-destructive' : ''}>
                        <SelectValue placeholder="Select estimated duration" />
                      </SelectTrigger>
                      <SelectContent>
                        {DURATIONS.map(d => (
                          <SelectItem key={d.value} value={d.value}>{d.label}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    {errors.duration && (
                      <p className="text-xs text-destructive flex items-center gap-1">
                        <AlertCircle className="w-3 h-3" /> {errors.duration}
                      </p>
                    )}
                  </div>

                  {/* Experience Level */}
                  <div className="space-y-2">
                    <label className="text-sm font-medium">
                      Experience Level <span className="text-destructive">*</span>
                    </label>
                    <Select value={experienceLevel} onValueChange={setExperienceLevel}>
                      <SelectTrigger className={errors.experienceLevel ? 'border-destructive' : ''}>
                        <SelectValue placeholder="Select experience level" />
                      </SelectTrigger>
                      <SelectContent>
                        {EXPERIENCE_LEVELS.map(e => (
                          <SelectItem key={e.value} value={e.value}>{e.label}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                    {errors.experienceLevel && (
                      <p className="text-xs text-destructive flex items-center gap-1">
                        <AlertCircle className="w-3 h-3" /> {errors.experienceLevel}
                      </p>
                    )}
                  </div>

                  {/* Urgent Switch */}
                  <div className="flex items-center justify-between p-4 rounded-xl bg-amber-50 border border-amber-200">
                    <div>
                      <div className="font-medium text-sm text-amber-900">Mark as Urgent</div>
                      <div className="text-xs text-amber-700 mt-0.5">
                        This will highlight your job to attract freelancers faster
                      </div>
                    </div>
                    <Switch checked={isUrgent} onCheckedChange={setIsUrgent} />
                  </div>
                </CardContent>
              </Card>
            </motion.div>
          )}

          {/* Step 3: Review */}
          {step === 3 && (
            <motion.div
              key="step3"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
            >
              <Card className="border-0 shadow-sm">
                <CardHeader>
                  <CardTitle className="text-xl">Review Your Job Posting</CardTitle>
                </CardHeader>
                <CardContent className="space-y-6">
                  <div className="p-6 rounded-xl bg-gray-50 space-y-4">
                    <div>
                      <h3 className="text-lg font-semibold text-gray-900">{title || 'Untitled Job'}</h3>
                      <Badge className="mt-2 bg-emerald-50 text-emerald-700">
                        {category ? CATEGORIES.find(c => c.value === category)?.label : 'No category'}
                      </Badge>
                    </div>

                    <Separator />

                    <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
                      <div>
                        <div className="text-xs text-muted-foreground">Budget Type</div>
                        <div className="text-sm font-medium">{budgetType === 'FIXED_PRICE' ? 'Fixed Price' : 'Hourly'}</div>
                      </div>
                      <div>
                        <div className="text-xs text-muted-foreground">
                          {budgetType === 'FIXED_PRICE' ? 'Budget' : 'Rate Range'}
                        </div>
                        <div className="text-sm font-medium">
                          ${budgetMin || '0'}{budgetMax ? ` - $${budgetMax}` : ''}
                          {budgetType === 'HOURLY' ? '/hr' : ''}
                        </div>
                      </div>
                      <div>
                        <div className="text-xs text-muted-foreground">Duration</div>
                        <div className="text-sm font-medium">
                          {DURATIONS.find(d => d.value === duration)?.label || 'Not set'}
                        </div>
                      </div>
                      <div>
                        <div className="text-xs text-muted-foreground">Experience Level</div>
                        <div className="text-sm font-medium">
                          {EXPERIENCE_LEVELS.find(e => e.value === experienceLevel)?.label || 'Not set'}
                        </div>
                      </div>
                      <div>
                        <div className="text-xs text-muted-foreground">Urgent</div>
                        <div className="text-sm font-medium">{isUrgent ? 'Yes' : 'No'}</div>
                      </div>
                    </div>

                    {skills.length > 0 && (
                      <div>
                        <div className="text-xs text-muted-foreground mb-2">Skills</div>
                        <div className="flex flex-wrap gap-1.5">
                          {skills.map(skill => (
                            <Badge key={skill} variant="outline" className="font-normal">{skill}</Badge>
                          ))}
                        </div>
                      </div>
                    )}

                    <Separator />

                    <div>
                      <div className="text-xs text-muted-foreground mb-2">Description</div>
                      <div className="text-sm text-muted-foreground leading-relaxed whitespace-pre-wrap max-h-48 overflow-y-auto custom-scrollbar">
                        {description}
                      </div>
                    </div>
                  </div>

                  {/* Edit links */}
                  <div className="flex gap-4">
                    <button
                      onClick={() => setStep(1)}
                      className="text-sm text-emerald-600 hover:text-emerald-700 font-medium"
                    >
                      Edit Job Details
                    </button>
                    <button
                      onClick={() => setStep(2)}
                      className="text-sm text-emerald-600 hover:text-emerald-700 font-medium"
                    >
                      Edit Budget & Timeline
                    </button>
                  </div>
                </CardContent>
              </Card>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Navigation Buttons */}
        <div className="flex items-center justify-between mt-6">
          {step > 1 ? (
            <Button variant="outline" onClick={prevStep}>
              <ChevronLeft className="w-4 h-4 mr-2" />
              Previous
            </Button>
          ) : (
            <Button variant="ghost" onClick={goBack}>
              Cancel
            </Button>
          )}

          {step < 3 ? (
            <Button className="bg-emerald-600 hover:bg-emerald-700 text-white" onClick={nextStep}>
              Next
              <ChevronRight className="w-4 h-4 ml-2" />
            </Button>
          ) : (
            <div className="flex gap-3">
              <Button
                variant="outline"
                onClick={() => submitJob(true)}
                disabled={submitting}
              >
                Save as Draft
              </Button>
              <Button
                className="bg-emerald-600 hover:bg-emerald-700 text-white"
                onClick={() => submitJob(false)}
                disabled={submitting}
              >
                {submitting ? 'Publishing...' : 'Publish Job'}
              </Button>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
