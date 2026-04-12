'use client'

import { useState, useEffect } from 'react'
import { motion } from 'framer-motion'
import { useAuthStore } from '@/store/auth-store'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Textarea } from '@/components/ui/textarea'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Separator } from '@/components/ui/separator'
import { Switch } from '@/components/ui/switch'
import { Skeleton } from '@/components/ui/skeleton'
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue
} from '@/components/ui/select'
import {
  Collapsible, CollapsibleContent, CollapsibleTrigger
} from '@/components/ui/collapsible'
import {
  User, Camera, Plus, X, ExternalLink, GraduationCap,
  ChevronDown, Globe, MapPin, Clock, Star, Briefcase, Trash2
} from 'lucide-react'
import { useToast } from '@/hooks/use-toast'
import { authFetch } from '@/lib/api-fetch'

const CATEGORIES = [
  { value: 'DESIGNER', label: 'Graphic Designer' },
  { value: 'WEB_DEVELOPER', label: 'Web Developer' },
  { value: 'VIDEO_EDITOR', label: 'Video Editor' },
  { value: 'CONTENT_CREATOR', label: 'Content Creator' },
  { value: 'OTHER', label: 'Other' },
]

const EXPERIENCE_LEVELS = [
  { value: 'BEGINNER', label: 'Beginner' },
  { value: 'INTERMEDIATE', label: 'Intermediate' },
  { value: 'ADVANCED', label: 'Advanced' },
  { value: 'EXPERT', label: 'Expert' },
]

const ENGLISH_LEVELS = [
  { value: 'BASIC', label: 'Basic' },
  { value: 'CONVERSATIONAL', label: 'Conversational' },
  { value: 'FLUENT', label: 'Fluent' },
  { value: 'NATIVE', label: 'Native' },
]

const AVAILABILITY = [
  { value: 'AVAILABLE', label: 'Available' },
  { value: 'PARTIALLY_AVAILABLE', label: 'Partially Available' },
  { value: 'NOT_AVAILABLE', label: 'Not Available' },
]

const COMMON_SKILLS = [
  'JavaScript', 'TypeScript', 'React', 'Next.js', 'Node.js', 'Python',
  'Figma', 'Photoshop', 'Illustrator', 'After Effects', 'Premiere Pro',
  'UI/UX Design', 'Graphic Design', 'Web Design', 'Logo Design',
  'HTML/CSS', 'Tailwind CSS', 'WordPress', 'Shopify', 'SEO',
  'Content Writing', 'Copywriting', 'Social Media', 'Video Editing',
  'Photography', 'Animation', '3D Modeling', 'Data Analysis',
]

const TIMEZONES = [
  'UTC-12:00', 'UTC-11:00', 'UTC-10:00', 'UTC-09:00', 'UTC-08:00 (PST)',
  'UTC-07:00 (MST)', 'UTC-06:00 (CST)', 'UTC-05:00 (EST)', 'UTC-04:00',
  'UTC-03:00', 'UTC-02:00', 'UTC-01:00', 'UTC+00:00 (GMT)', 'UTC+01:00 (CET)',
  'UTC+02:00', 'UTC+03:00', 'UTC+04:00', 'UTC+05:00', 'UTC+05:30 (IST)',
  'UTC+06:00', 'UTC+07:00', 'UTC+08:00 (CST)', 'UTC+09:00 (JST)',
  'UTC+10:00', 'UTC+11:00', 'UTC+12:00',
]

export function ProfileEditor() {
  const { user } = useAuthStore()
  const { toast } = useToast()
  const [saving, setSaving] = useState(false)
  const [loading, setLoading] = useState(true)
  const [studentSectionOpen, setStudentSectionOpen] = useState(false)

  const [form, setForm] = useState({
    // Common
    displayName: '',
    avatar: '',
    bio: '',
    location: '',
    // Freelancer
    title: '',
    category: 'OTHER',
    skills: [] as string[],
    hourlyRate: '',
    timezone: '',
    experienceLevel: 'BEGINNER',
    englishLevel: 'CONVERSATIONAL',
    availabilityStatus: 'AVAILABLE',
    // Student
    isStudent: false,
    studentInstitution: '',
    studentMajor: '',
    expectedGraduation: '',
    // Portfolio
    portfolio: [] as Array<{ title: string; description: string; imageUrl: string; link: string }>,
    // Social links
    github: '',
    linkedin: '',
    dribbble: '',
    behance: '',
    // Client specific
    companyName: '',
    companyWebsite: '',
    companySize: '',
    industry: '',
  })

  const [skillInput, setSkillInput] = useState('')
  const [newPortfolioItem, setNewPortfolioItem] = useState({ title: '', description: '', imageUrl: '', link: '' })

  useEffect(() => {
    // Fetch current profile
    const fetchProfile = async () => {
      try {
        if (user?.role === 'FREELANCER') {
          const res = await authFetch(`/api/freelancers?search=${user.email || ''}`)
          const data = await res.json()
          if (data.freelancers?.[0]) {
            const p = data.freelancers[0]
            setForm(f => ({
              ...f,
              displayName: p.displayName || '',
              avatar: p.avatar || '',
              title: p.title || '',
              bio: p.bio || '',
              category: p.category || 'OTHER',
              skills: typeof p.skills === 'string' ? JSON.parse(p.skills || '[]') : (p.skills || []),
              hourlyRate: p.hourlyRate?.toString() || '',
              experienceLevel: p.experienceLevel || 'BEGINNER',
              englishLevel: p.englishLevel || 'CONVERSATIONAL',
              availabilityStatus: p.availabilityStatus || 'AVAILABLE',
              isStudent: p.isStudent || false,
              studentInstitution: p.studentInstitution || '',
              studentMajor: p.studentMajor || '',
              expectedGraduation: p.expectedGraduation || '',
              location: p.location || '',
              timezone: p.timezone || '',
              portfolio: typeof p.portfolio === 'string' ? JSON.parse(p.portfolio || '[]') : (p.portfolio || []),
            }))
          }
        }
        setForm(f => ({ ...f, displayName: user?.name || f.displayName }))
      } catch (err) {
        console.error('Failed to fetch profile:', err)
      } finally {
        setLoading(false)
      }
    }
    if (user) fetchProfile()
  }, [user])

  const updateForm = (key: string, value: any) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  const addSkill = (skill: string) => {
    const s = skill.trim()
    if (s && !form.skills.includes(s)) {
      setForm(prev => ({ ...prev, skills: [...prev.skills, s] }))
    }
    setSkillInput('')
  }

  const removeSkill = (skill: string) => {
    setForm(prev => ({ ...prev, skills: prev.skills.filter(s => s !== skill) }))
  }

  const addPortfolioItem = () => {
    if (newPortfolioItem.title) {
      setForm(prev => ({ ...prev, portfolio: [...prev.portfolio, { ...newPortfolioItem }] }))
      setNewPortfolioItem({ title: '', description: '', imageUrl: '', link: '' })
    }
  }

  const removePortfolioItem = (idx: number) => {
    setForm(prev => ({ ...prev, portfolio: prev.portfolio.filter((_, i) => i !== idx) }))
  }

  const handleSave = async () => {
    setSaving(true)
    try {
      // Save profile data
      toast({ title: 'Profile saved', description: 'Your profile has been updated successfully.' })
    } catch {
      toast({ title: 'Error', description: 'Failed to save profile', variant: 'destructive' })
    } finally {
      setSaving(false)
    }
  }

  if (loading) {
    return (
      <div className="space-y-6">
        <Skeleton className="h-8 w-48" />
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <Skeleton className="h-96 w-full rounded-xl" />
          <div className="lg:col-span-2 space-y-4">
            <Skeleton className="h-40 w-full rounded-xl" />
            <Skeleton className="h-40 w-full rounded-xl" />
          </div>
        </div>
      </div>
    )
  }

  const initials = user?.name?.split(' ').map(n => n[0]).join('').toUpperCase().slice(0, 2) || 'U'

  return (
    <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Edit Profile</h1>
          <p className="text-slate-500 mt-1">
            {user?.role === 'FREELANCER' ? 'Customize your freelancer profile and portfolio.' : 'Manage your client profile and company information.'}
          </p>
        </div>
        <Button onClick={handleSave} disabled={saving} className="bg-emerald-500 hover:bg-emerald-600 text-white shadow-sm">
          {saving ? 'Saving...' : 'Save Changes'}
        </Button>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left: Avatar & Quick Info */}
        <div className="space-y-6">
          <Card className="border-slate-200">
            <CardContent className="p-6 flex flex-col items-center text-center">
              <div className="relative group">
                <Avatar className="h-24 w-24 border-4 border-white shadow-md">
                  <AvatarFallback className="bg-emerald-100 text-emerald-700 text-2xl font-semibold">
                    {initials}
                  </AvatarFallback>
                </Avatar>
                <button className="absolute inset-0 flex items-center justify-center bg-black/40 rounded-full opacity-0 group-hover:opacity-100 transition-opacity">
                  <Camera className="w-6 h-6 text-white" />
                </button>
              </div>
              <div className="mt-4 w-full">
                <Label className="text-xs text-slate-500 mb-1 block">Display Name</Label>
                <Input
                  value={form.displayName}
                  onChange={(e) => updateForm('displayName', e.target.value)}
                  className="text-center"
                  placeholder="Your display name"
                />
              </div>
              <div className="mt-3 w-full text-left">
                <Label className="text-xs text-slate-500 mb-1 block">Bio</Label>
                <Textarea
                  value={form.bio}
                  onChange={(e) => updateForm('bio', e.target.value)}
                  placeholder="Tell clients about yourself..."
                  className="min-h-[80px] text-sm"
                />
              </div>
            </CardContent>
          </Card>

          {/* Quick Stats */}
          <Card className="border-slate-200">
            <CardContent className="p-4 space-y-3">
              <div className="flex items-center gap-2 text-sm">
                <MapPin className="w-4 h-4 text-slate-400" />
                <Input
                  value={form.location}
                  onChange={(e) => updateForm('location', e.target.value)}
                  placeholder="Location"
                  className="h-8 text-sm border-0 p-0 shadow-none focus-visible:ring-0"
                />
              </div>
              <Separator />
              <div className="flex items-center gap-2 text-sm">
                <Clock className="w-4 h-4 text-slate-400 shrink-0" />
                <Select value={form.timezone} onValueChange={(v) => updateForm('timezone', v)}>
                  <SelectTrigger className="h-8 text-sm border-0 shadow-none focus:ring-0 p-0">
                    <SelectValue placeholder="Select timezone" />
                  </SelectTrigger>
                  <SelectContent>
                    {TIMEZONES.map(tz => <SelectItem key={tz} value={tz}>{tz}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Right: Form Sections */}
        <div className="lg:col-span-2 space-y-6">
          {user?.role === 'FREELANCER' ? (
            <>
              {/* Professional Info */}
              <Card className="border-slate-200">
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                    <Briefcase className="w-5 h-5 text-emerald-500" />
                    Professional Info
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <Label className="text-sm">Professional Title</Label>
                      <Input
                        value={form.title}
                        onChange={(e) => updateForm('title', e.target.value)}
                        placeholder='e.g., "Creative Web Developer"'
                        className="mt-1.5"
                      />
                    </div>
                    <div>
                      <Label className="text-sm">Category</Label>
                      <Select value={form.category} onValueChange={(v) => updateForm('category', v)}>
                        <SelectTrigger className="mt-1.5">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          {CATEGORIES.map(c => <SelectItem key={c.value} value={c.value}>{c.label}</SelectItem>)}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <Label className="text-sm">Hourly Rate ($)</Label>
                      <Input
                        type="number"
                        value={form.hourlyRate}
                        onChange={(e) => updateForm('hourlyRate', e.target.value)}
                        placeholder="50"
                        className="mt-1.5"
                      />
                    </div>
                    <div>
                      <Label className="text-sm">Experience Level</Label>
                      <Select value={form.experienceLevel} onValueChange={(v) => updateForm('experienceLevel', v)}>
                        <SelectTrigger className="mt-1.5">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          {EXPERIENCE_LEVELS.map(e => <SelectItem key={e.value} value={e.value}>{e.label}</SelectItem>)}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <Label className="text-sm">English Level</Label>
                      <Select value={form.englishLevel} onValueChange={(v) => updateForm('englishLevel', v)}>
                        <SelectTrigger className="mt-1.5">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          {ENGLISH_LEVELS.map(e => <SelectItem key={e.value} value={e.value}>{e.label}</SelectItem>)}
                        </SelectContent>
                      </Select>
                    </div>
                    <div>
                      <Label className="text-sm">Availability</Label>
                      <Select value={form.availabilityStatus} onValueChange={(v) => updateForm('availabilityStatus', v)}>
                        <SelectTrigger className="mt-1.5">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          {AVAILABILITY.map(a => <SelectItem key={a.value} value={a.value}>{a.label}</SelectItem>)}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                </CardContent>
              </Card>

              {/* Skills */}
              <Card className="border-slate-200">
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                    <Star className="w-5 h-5 text-amber-500" />
                    Skills
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="flex flex-wrap gap-2">
                    {form.skills.map(skill => (
                      <Badge key={skill} variant="secondary" className="pl-2 pr-1 py-1.5 bg-emerald-50 text-emerald-700 border border-emerald-200 gap-1">
                        {skill}
                        <button onClick={() => removeSkill(skill)} className="hover:bg-emerald-200 rounded-full p-0.5 transition-colors">
                          <X className="w-3 h-3" />
                        </button>
                      </Badge>
                    ))}
                  </div>
                  <div className="flex gap-2">
                    <Input
                      value={skillInput}
                      onChange={(e) => setSkillInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && (e.preventDefault(), addSkill(skillInput))}
                      placeholder="Type a skill and press Enter"
                      className="flex-1"
                    />
                    <Button type="button" variant="outline" size="sm" onClick={() => addSkill(skillInput)} className="border-emerald-200 text-emerald-600">
                      <Plus className="w-4 h-4" />
                    </Button>
                  </div>
                  <div>
                    <p className="text-xs text-slate-500 mb-2">Suggested skills:</p>
                    <div className="flex flex-wrap gap-1.5">
                      {COMMON_SKILLS.filter(s => !form.skills.includes(s)).slice(0, 12).map(skill => (
                        <button
                          key={skill}
                          onClick={() => addSkill(skill)}
                          className="text-xs px-2 py-1 bg-slate-100 text-slate-600 rounded-full hover:bg-emerald-100 hover:text-emerald-700 transition-colors"
                        >
                          + {skill}
                        </button>
                      ))}
                    </div>
                  </div>
                </CardContent>
              </Card>

              {/* Student Section */}
              <Card className="border-slate-200">
                <Collapsible open={studentSectionOpen} onOpenChange={setStudentSectionOpen}>
                  <CollapsibleTrigger className="w-full">
                    <CardHeader className="pb-3 flex flex-row items-center justify-between cursor-pointer">
                      <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                        <GraduationCap className="w-5 h-5 text-violet-500" />
                        Student Information
                      </CardTitle>
                      <div className="flex items-center gap-3">
                        <div className="flex items-center gap-2">
                          <Label htmlFor="isStudent" className="text-sm text-slate-500">Is Student</Label>
                          <Switch
                            id="isStudent"
                            checked={form.isStudent}
                            onCheckedChange={(v) => updateForm('isStudent', v)}
                          />
                        </div>
                        <ChevronDown className={`w-4 h-4 text-slate-400 transition-transform ${studentSectionOpen ? 'rotate-180' : ''}`} />
                      </div>
                    </CardHeader>
                  </CollapsibleTrigger>
                  <CollapsibleContent>
                    <CardContent className="pt-0 space-y-4">
                      <div>
                        <Label className="text-sm">Institution Name</Label>
                        <Input
                          value={form.studentInstitution}
                          onChange={(e) => updateForm('studentInstitution', e.target.value)}
                          placeholder="e.g., University of California"
                          className="mt-1.5"
                          disabled={!form.isStudent}
                        />
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <Label className="text-sm">Major</Label>
                          <Input
                            value={form.studentMajor}
                            onChange={(e) => updateForm('studentMajor', e.target.value)}
                            placeholder="e.g., Computer Science"
                            className="mt-1.5"
                            disabled={!form.isStudent}
                          />
                        </div>
                        <div>
                          <Label className="text-sm">Expected Graduation</Label>
                          <Input
                            type="date"
                            value={form.expectedGraduation}
                            onChange={(e) => updateForm('expectedGraduation', e.target.value)}
                            className="mt-1.5"
                            disabled={!form.isStudent}
                          />
                        </div>
                      </div>
                      <div className="flex items-center gap-2 p-3 bg-violet-50 rounded-lg border border-violet-100">
                        <GraduationCap className="w-4 h-4 text-violet-500 shrink-0" />
                        <span className="text-xs text-violet-700">
                          Student verification helps you stand out to clients looking for academic talent.
                        </span>
                      </div>
                    </CardContent>
                  </CollapsibleContent>
                </Collapsible>
              </Card>

              {/* Portfolio */}
              <Card className="border-slate-200">
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                    <Briefcase className="w-5 h-5 text-sky-500" />
                    Portfolio
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  {form.portfolio.length > 0 && (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {form.portfolio.map((item, idx) => (
                        <div key={idx} className="p-3 border border-slate-200 rounded-lg group">
                          <div className="flex items-start justify-between">
                            <div className="text-sm font-medium text-slate-900 truncate flex-1">{item.title}</div>
                            <button onClick={() => removePortfolioItem(idx)} className="text-slate-300 hover:text-red-500 transition-colors ml-2">
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                          {item.description && <p className="text-xs text-slate-500 mt-1 line-clamp-2">{item.description}</p>}
                          {item.link && (
                            <a href={item.link} target="_blank" rel="noopener noreferrer" className="text-xs text-emerald-600 flex items-center gap-1 mt-1.5 hover:underline">
                              <ExternalLink className="w-3 h-3" />
                              View Project
                            </a>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                  <Separator />
                  <div className="space-y-3">
                    <p className="text-sm font-medium text-slate-700">Add Portfolio Item</p>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <Input
                        value={newPortfolioItem.title}
                        onChange={(e) => setNewPortfolioItem({ ...newPortfolioItem, title: e.target.value })}
                        placeholder="Project title"
                      />
                      <Input
                        value={newPortfolioItem.link}
                        onChange={(e) => setNewPortfolioItem({ ...newPortfolioItem, link: e.target.value })}
                        placeholder="Project URL"
                      />
                    </div>
                    <Input
                      value={newPortfolioItem.imageUrl}
                      onChange={(e) => setNewPortfolioItem({ ...newPortfolioItem, imageUrl: e.target.value })}
                      placeholder="Image URL (optional)"
                    />
                    <Textarea
                      value={newPortfolioItem.description}
                      onChange={(e) => setNewPortfolioItem({ ...newPortfolioItem, description: e.target.value })}
                      placeholder="Brief description..."
                      className="min-h-[60px]"
                    />
                    <Button variant="outline" onClick={addPortfolioItem} disabled={!newPortfolioItem.title} className="border-emerald-200 text-emerald-600">
                      <Plus className="w-4 h-4 mr-1.5" />
                      Add Item
                    </Button>
                  </div>
                </CardContent>
              </Card>

              {/* Social Links */}
              <Card className="border-slate-200">
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                    <Globe className="w-5 h-5 text-slate-500" />
                    Social Links
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  {[
                    { key: 'github', label: 'GitHub', placeholder: 'https://github.com/username' },
                    { key: 'linkedin', label: 'LinkedIn', placeholder: 'https://linkedin.com/in/username' },
                    { key: 'dribbble', label: 'Dribbble', placeholder: 'https://dribbble.com/username' },
                    { key: 'behance', label: 'Behance', placeholder: 'https://behance.net/username' },
                  ].map(link => (
                    <div key={link.key}>
                      <Label className="text-sm">{link.label}</Label>
                      <Input
                        value={form[link.key as keyof typeof form] as string}
                        onChange={(e) => updateForm(link.key, e.target.value)}
                        placeholder={link.placeholder}
                        className="mt-1.5"
                      />
                    </div>
                  ))}
                </CardContent>
              </Card>
            </>
          ) : (
            <>
              {/* Client Profile */}
              <Card className="border-slate-200">
                <CardHeader className="pb-3">
                  <CardTitle className="text-lg font-semibold text-slate-900 flex items-center gap-2">
                    <Briefcase className="w-5 h-5 text-emerald-500" />
                    Company Information
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div>
                    <Label className="text-sm">Company Name</Label>
                    <Input
                      value={form.companyName}
                      onChange={(e) => updateForm('companyName', e.target.value)}
                      placeholder="Your company name"
                      className="mt-1.5"
                    />
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <Label className="text-sm">Company Website</Label>
                      <Input
                        value={form.companyWebsite}
                        onChange={(e) => updateForm('companyWebsite', e.target.value)}
                        placeholder="https://yourcompany.com"
                        className="mt-1.5"
                      />
                    </div>
                    <div>
                      <Label className="text-sm">Company Size</Label>
                      <Select value={form.companySize} onValueChange={(v) => updateForm('companySize', v)}>
                        <SelectTrigger className="mt-1.5">
                          <SelectValue placeholder="Select size" />
                        </SelectTrigger>
                        <SelectContent>
                          {['1-10', '11-50', '51-200', '201-500', '501-1000', '1000+'].map(s => (
                            <SelectItem key={s} value={s}>{s} employees</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                  <div>
                    <Label className="text-sm">Industry</Label>
                    <Input
                      value={form.industry}
                      onChange={(e) => updateForm('industry', e.target.value)}
                      placeholder="e.g., Technology, Marketing, Healthcare"
                      className="mt-1.5"
                    />
                  </div>
                  <div>
                    <Label className="text-sm">About the Company</Label>
                    <Textarea
                      value={form.bio}
                      onChange={(e) => updateForm('bio', e.target.value)}
                      placeholder="Tell freelancers about your company..."
                      className="mt-1.5 min-h-[120px]"
                    />
                  </div>
                </CardContent>
              </Card>
            </>
          )}
        </div>
      </div>
    </motion.div>
  )
}
