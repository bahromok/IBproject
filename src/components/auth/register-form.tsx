'use client'

import { useState } from 'react'
import { motion, AnimatePresence } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Label } from '@/components/ui/label'
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card'
import { Checkbox } from '@/components/ui/checkbox'
import { Badge } from '@/components/ui/badge'
import { useAuthStore } from '@/store/auth-store'
import { useNavigationStore } from '@/store/navigation-store'
import { useToast } from '@/hooks/use-toast'
import { authClient } from '@/lib/auth/client'
import {
  GraduationCap,
  Eye,
  EyeOff,
  Loader2,
  ArrowLeft,
  ArrowRight,
  Check,
  Briefcase,
  Code,
  Palette,
  Film,
  FileText,
  Users,
  DollarSign,
  School,
} from 'lucide-react'

const FREELANCER_CATEGORIES = [
  { value: 'DESIGNER', label: 'Designer', icon: Palette, color: 'text-pink-500' },
  { value: 'WEB_DEVELOPER', label: 'Web Developer', icon: Code, color: 'text-emerald-500' },
  { value: 'VIDEO_EDITOR', label: 'Video Editor', icon: Film, color: 'text-purple-500' },
  { value: 'CONTENT_CREATOR', label: 'Content Creator', icon: FileText, color: 'text-amber-500' },
]

const slideVariants = {
  enter: (direction: number) => ({
    x: direction > 0 ? 100 : -100,
    opacity: 0,
  }),
  center: {
    x: 0,
    opacity: 1,
  },
  exit: (direction: number) => ({
    x: direction > 0 ? -100 : 100,
    opacity: 0,
  }),
}

export function RegisterForm() {
  const [step, setStep] = useState(1)
  const [direction, setDirection] = useState(1)
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [agreedToTerms, setAgreedToTerms] = useState(false)
  const [showPassword, setShowPassword] = useState(false)
  const { register } = useAuthStore()
  const { navigate } = useNavigationStore()
  const { toast } = useToast()

  // Form data
  const [role, setRole] = useState<'CLIENT' | 'FREELANCER' | ''>('')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [displayName, setDisplayName] = useState('')
  const [companyName, setCompanyName] = useState('')
  const [industry, setIndustry] = useState('')
  const [category, setCategory] = useState('')
  const [skillInput, setSkillInput] = useState('')
  const [skills, setSkills] = useState<string[]>([])
  const [hourlyRate, setHourlyRate] = useState('')
  const [isStudent, setIsStudent] = useState(false)
  const [institution, setInstitution] = useState('')
  const [bio, setBio] = useState('')
  const [verificationCode, setVerificationCode] = useState('')
  const [isVerifying, setIsVerifying] = useState(false)

  const totalSteps = step === 4 ? 4 : (role === 'FREELANCER' ? 3 : (role === 'CLIENT' ? 3 : 1))

  const goNext = () => {
    setDirection(1)
    setStep((s) => Math.min(s + 1, totalSteps))
  }

  const goBack = () => {
    if (step === 1) {
      navigate('auth/login')
      return
    }
    setDirection(-1)
    setStep((s) => Math.max(s - 1, 1))
  }

  const addSkill = () => {
    const trimmed = skillInput.trim()
    if (trimmed && !skills.includes(trimmed) && skills.length < 10) {
      setSkills([...skills, trimmed])
      setSkillInput('')
    }
  }

  const removeSkill = (skill: string) => {
    setSkills(skills.filter((s) => s !== skill))
  }

  const handleSkillKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.preventDefault()
      addSkill()
    }
  }

  const validateStep = (): boolean => {
    if (step === 1) {
      if (!role) {
        toast({ title: 'Select a role', description: 'Please choose whether you want to hire or work.', variant: 'destructive' })
        return false
      }
      return true
    }

    if (step === 2) {
      if (!email.trim() || !email.includes('@')) {
        toast({ title: 'Invalid email', description: 'Please enter a valid email address.', variant: 'destructive' })
        return false
      }
      if (password.length < 8) {
        toast({ title: 'Weak password', description: 'Password must be at least 8 characters.', variant: 'destructive' })
        return false
      }
      if (password !== confirmPassword) {
        toast({ title: 'Passwords don\'t match', description: 'Please make sure both passwords match.', variant: 'destructive' })
        return false
      }
      return true
    }

    if (step === 3) {
      if (!displayName.trim()) {
        toast({ title: 'Name required', description: 'Please enter your display name.', variant: 'destructive' })
        return false
      }
      if (role === 'FREELANCER' && !category) {
        toast({ title: 'Category required', description: 'Please select your primary category.', variant: 'destructive' })
        return false
      }
      if (!agreedToTerms) {
        toast({ title: 'Terms required', description: 'You must agree to the Terms of Service.', variant: 'destructive' })
        return false
      }
      return true
    }

    return true
  }

  const handleSubmit = async () => {
    if (!validateStep()) return

    setIsSubmitting(true)

    const result = await register({
      email,
      password,
      role: role as 'CLIENT' | 'FREELANCER',
      displayName: displayName.trim(),
      companyName: companyName.trim() || undefined,
      industry: industry.trim() || undefined,
      category: category || undefined,
      skills: skills.length > 0 ? skills : undefined,
      hourlyRate: hourlyRate ? parseFloat(hourlyRate) : undefined,
      isStudent,
      institution: institution.trim() || undefined,
      bio: bio.trim() || undefined,
    })

    if (result.success) {
      toast({ title: 'Code Sent', description: 'Please check your email for the verification code.' })
      setStep(4)
    } else {
      toast({ title: 'Registration failed', description: result.error || 'Please try again.', variant: 'destructive' })
    }
    setIsSubmitting(false)
  }

  const handleVerify = async () => {
    if (!verificationCode) {
      toast({ title: 'Error', description: 'Please enter the verification code.', variant: 'destructive' })
      return
    }

    setIsVerifying(true)
    try {
      const { error } = await authClient.emailOtp.verifyEmail({
        email,
        otp: verificationCode.trim(),
      })

      if (error) {
        toast({ title: 'Verification Failed', description: error.message || 'Invalid code', variant: 'destructive' })
      } else {
        toast({ title: 'Verified!', description: 'Your email has been verified successfully.' })
        navigate('dashboard')
      }
    } catch {
      toast({ title: 'Error', description: 'Verification failed', variant: 'destructive' })
    } finally {
      setIsVerifying(false)
    }
  }

  return (
    <div className="min-h-screen flex flex-col bg-white">
      {/* Header */}
      <header className="border-b bg-white sticky top-0 z-10">
        <div className="max-w-3xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center">
          <button
            onClick={() => navigate('landing')}
            className="flex items-center gap-2"
          >
            <div className="w-8 h-8 rounded-lg bg-emerald-500 flex items-center justify-center">
              <GraduationCap className="w-4 h-4 text-white" />
            </div>
            <span className="text-lg font-bold text-slate-900">
              Student<span className="text-emerald-500">Hire</span>
            </span>
          </button>
        </div>
      </header>

      {/* Progress Bar */}
      <div className="bg-slate-50 border-b">
        <div className="max-w-3xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex items-center justify-between mb-3">
            <button
              onClick={goBack}
              className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-700 transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              Back
            </button>
            <span className="text-sm font-medium text-slate-600">
              Step {step} of {totalSteps}
            </span>
          </div>

          <div className="flex gap-2">
            {Array.from({ length: totalSteps }).map((_, i) => (
              <div
                key={i}
                className={`h-1.5 flex-1 rounded-full transition-all duration-300 ${i < step ? 'bg-emerald-500' : i === step ? 'bg-emerald-300' : 'bg-slate-200'
                  }`}
              />
            ))}
          </div>

          <div className="mt-3 flex items-center justify-between text-xs text-slate-500">
            <span className={step >= 1 ? 'text-emerald-600 font-medium' : ''}>Choose Role</span>
            <span className={step >= 2 ? 'text-emerald-600 font-medium' : ''}>Account Info</span>
            <span className={step >= 3 ? 'text-emerald-600 font-medium' : ''}>Profile Details</span>
          </div>
        </div>
      </div>

      {/* Form Content */}
      <div className="flex-1 flex items-start justify-center px-4 sm:px-6 lg:px-8 py-8">
        <div className="w-full max-w-xl">
          <AnimatePresence mode="wait" custom={direction}>
            {/* Step 4: Email Verification */}
            {step === 4 && (
              <motion.div
                key="step4"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                className="space-y-6"
              >
                <div className="text-center space-y-2">
                  <div className="w-16 h-16 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-4">
                    <Check className="w-8 h-8 text-emerald-600" />
                  </div>
                  <h2 className="text-2xl font-bold text-slate-900">Verify your email</h2>
                  <p className="text-slate-500">
                    We&apos;ve sent a verification code to <span className="font-semibold text-slate-900">{email}</span>.
                  </p>
                </div>

                <div className="space-y-4">
                  <div className="space-y-2">
                    <Label htmlFor="otp" className="text-slate-700 font-medium">
                      Verification Code
                    </Label>
                    <Input
                      id="otp"
                      placeholder="Enter 8-character code"
                      value={verificationCode}
                      onChange={(e) => setVerificationCode(e.target.value)}
                      className="h-12 text-center text-xl font-bold tracking-widest border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                      disabled={isVerifying}
                    />
                  </div>

                  <Button
                    onClick={handleVerify}
                    className="w-full h-12 bg-emerald-500 hover:bg-emerald-600 text-white font-semibold shadow-lg shadow-emerald-500/25"
                    disabled={isVerifying}
                  >
                    {isVerifying ? (
                      <>
                        <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                        Verifying...
                      </>
                    ) : (
                      'Verify Account'
                    )}
                  </Button>

                  <div className="text-center">
                    <button
                      onClick={() => setStep(2)}
                      className="text-sm text-slate-500 hover:text-slate-700"
                    >
                      Didn&apos;t receive it? Check your email or <span className="text-emerald-600 font-medium">go back</span>
                    </button>
                  </div>
                </div>
              </motion.div>
            )}

            {/* Step 1: Choose Role */}
            {step === 1 && (
              <motion.div
                key="step1"
                custom={direction}
                variants={slideVariants}
                initial="enter"
                animate="center"
                exit="exit"
                transition={{ duration: 0.3 }}
              >
                <Card className="border-0 shadow-lg">
                  <CardHeader className="text-center pb-2">
                    <CardTitle className="text-2xl font-bold text-slate-900">
                      Join StudentHire
                    </CardTitle>
                    <CardDescription className="text-slate-500 mt-1">
                      How would you like to use our platform?
                    </CardDescription>
                  </CardHeader>
                  <CardContent className="pt-4 space-y-4">
                    <button
                      onClick={() => {
                        setRole('CLIENT')
                        setTimeout(goNext, 150)
                      }}
                      className={`w-full p-6 rounded-xl border-2 text-left transition-all duration-200 ${role === 'CLIENT'
                        ? 'border-emerald-500 bg-emerald-50 shadow-md'
                        : 'border-slate-200 hover:border-emerald-300 hover:shadow-sm'
                        }`}
                    >
                      <div className="flex items-start gap-4">
                        <div className={`w-12 h-12 rounded-xl flex items-center justify-center ${role === 'CLIENT' ? 'bg-emerald-500' : 'bg-slate-100'
                          }`}>
                          <Briefcase className={`w-6 h-6 ${role === 'CLIENT' ? 'text-white' : 'text-slate-500'}`} />
                        </div>
                        <div className="flex-1">
                          <h3 className="text-lg font-semibold text-slate-900">I want to hire</h3>
                          <p className="text-sm text-slate-500 mt-1">
                            Post jobs and find talented student freelancers for your projects.
                          </p>
                          <div className="flex gap-2 mt-3">
                            {['Post Jobs', 'Review Proposals', 'Manage Contracts'].map((tag) => (
                              <Badge key={tag} variant="secondary" className="text-xs bg-slate-100 text-slate-600">
                                {tag}
                              </Badge>
                            ))}
                          </div>
                        </div>
                        {role === 'CLIENT' && (
                          <div className="w-6 h-6 rounded-full bg-emerald-500 flex items-center justify-center shrink-0">
                            <Check className="w-4 h-4 text-white" />
                          </div>
                        )}
                      </div>
                    </button>

                    <button
                      onClick={() => {
                        setRole('FREELANCER')
                        setTimeout(goNext, 150)
                      }}
                      className={`w-full p-6 rounded-xl border-2 text-left transition-all duration-200 ${role === 'FREELANCER'
                        ? 'border-emerald-500 bg-emerald-50 shadow-md'
                        : 'border-slate-200 hover:border-emerald-300 hover:shadow-sm'
                        }`}
                    >
                      <div className="flex items-start gap-4">
                        <div className={`w-12 h-12 rounded-xl flex items-center justify-center ${role === 'FREELANCER' ? 'bg-emerald-500' : 'bg-slate-100'
                          }`}>
                          <GraduationCap className={`w-6 h-6 ${role === 'FREELANCER' ? 'text-white' : 'text-slate-500'}`} />
                        </div>
                        <div className="flex-1">
                          <h3 className="text-lg font-semibold text-slate-900">I want to work</h3>
                          <p className="text-sm text-slate-500 mt-1">
                            Find freelance opportunities and build your professional portfolio.
                          </p>
                          <div className="flex gap-2 mt-3">
                            {['Find Jobs', 'Submit Proposals', 'Earn Money'].map((tag) => (
                              <Badge key={tag} variant="secondary" className="text-xs bg-slate-100 text-slate-600">
                                {tag}
                              </Badge>
                            ))}
                          </div>
                        </div>
                        {role === 'FREELANCER' && (
                          <div className="w-6 h-6 rounded-full bg-emerald-500 flex items-center justify-center shrink-0">
                            <Check className="w-4 h-4 text-white" />
                          </div>
                        )}
                      </div>
                    </button>
                  </CardContent>
                </Card>
              </motion.div>
            )}

            {/* Step 2: Email + Password */}
            {step === 2 && (
              <motion.div
                key="step2"
                custom={direction}
                variants={slideVariants}
                initial="enter"
                animate="center"
                exit="exit"
                transition={{ duration: 0.3 }}
              >
                <Card className="border-0 shadow-lg">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-2xl font-bold text-slate-900">
                      Create your account
                    </CardTitle>
                    <CardDescription className="text-slate-500">
                      Enter your email and choose a secure password.
                    </CardDescription>
                  </CardHeader>
                  <CardContent className="pt-4 space-y-4">
                    <div className="space-y-2">
                      <Label htmlFor="reg-email" className="text-slate-700 font-medium">
                        Email address
                      </Label>
                      <Input
                        id="reg-email"
                        type="email"
                        placeholder="you@example.com"
                        value={email}
                        onChange={(e) => setEmail(e.target.value)}
                        autoComplete="email"
                        className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                      />
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="reg-password" className="text-slate-700 font-medium">
                        Password
                      </Label>
                      <div className="relative">
                        <Input
                          id="reg-password"
                          type={showPassword ? 'text' : 'password'}
                          placeholder="Minimum 8 characters"
                          value={password}
                          onChange={(e) => setPassword(e.target.value)}
                          autoComplete="new-password"
                          className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20 pr-10"
                        />
                        <button
                          type="button"
                          onClick={() => setShowPassword(!showPassword)}
                          className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                          aria-label={showPassword ? 'Hide password' : 'Show password'}
                        >
                          {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                        </button>
                      </div>
                      {password && (
                        <div className="flex gap-1">
                          {[1, 2, 3, 4].map((level) => (
                            <div
                              key={level}
                              className={`h-1 flex-1 rounded-full transition-colors ${password.length >= level * 3
                                ? level <= 1
                                  ? 'bg-red-400'
                                  : level <= 2
                                    ? 'bg-amber-400'
                                    : level <= 3
                                      ? 'bg-emerald-400'
                                      : 'bg-emerald-500'
                                : 'bg-slate-200'
                                }`}
                            />
                          ))}
                          <span className="text-xs text-slate-500 ml-2">
                            {password.length < 6 ? 'Weak' : password.length < 10 ? 'Good' : 'Strong'}
                          </span>
                        </div>
                      )}
                    </div>

                    <div className="space-y-2">
                      <Label htmlFor="reg-confirm-password" className="text-slate-700 font-medium">
                        Confirm password
                      </Label>
                      <Input
                        id="reg-confirm-password"
                        type="password"
                        placeholder="Re-enter your password"
                        value={confirmPassword}
                        onChange={(e) => setConfirmPassword(e.target.value)}
                        autoComplete="new-password"
                        className={`h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20 ${confirmPassword && confirmPassword !== password ? 'border-red-300 focus:border-red-500' : ''
                          }`}
                      />
                      {confirmPassword && confirmPassword !== password && (
                        <p className="text-xs text-red-500">Passwords do not match</p>
                      )}
                    </div>
                  </CardContent>
                  <CardFooter>
                    <Button
                      onClick={() => {
                        if (validateStep()) goNext()
                      }}
                      className="w-full h-11 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg font-semibold shadow-lg shadow-emerald-500/25"
                    >
                      Continue
                      <ArrowRight className="w-4 h-4 ml-1" />
                    </Button>
                  </CardFooter>
                </Card>
              </motion.div>
            )}

            {/* Step 3: Profile Details */}
            {step === 3 && (
              <motion.div
                key="step3"
                custom={direction}
                variants={slideVariants}
                initial="enter"
                animate="center"
                exit="exit"
                transition={{ duration: 0.3 }}
              >
                <Card className="border-0 shadow-lg">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-2xl font-bold text-slate-900">
                      Tell us about yourself
                    </CardTitle>
                    <CardDescription className="text-slate-500">
                      {role === 'CLIENT'
                        ? 'Help freelancers understand what you need.'
                        : 'Showcase your skills to attract clients.'}
                    </CardDescription>
                  </CardHeader>
                  <CardContent className="pt-4 space-y-5">
                    {/* Common fields */}
                    <div className="space-y-2">
                      <Label htmlFor="display-name" className="text-slate-700 font-medium">
                        Display Name
                      </Label>
                      <Input
                        id="display-name"
                        placeholder={role === 'CLIENT' ? 'Your business name or alias' : 'Your creative alias'}
                        value={displayName}
                        onChange={(e) => setDisplayName(e.target.value)}
                        className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                      />
                      <p className="text-xs text-slate-400">
                        This is how other users will see you on the platform. You can change this later.
                      </p>
                    </div>

                    {/* Client-specific fields */}
                    {role === 'CLIENT' && (
                      <>
                        <div className="space-y-2">
                          <Label htmlFor="company-name" className="text-slate-700 font-medium">
                            Company Name <span className="text-slate-400 font-normal">(optional)</span>
                          </Label>
                          <Input
                            id="company-name"
                            placeholder="Your company or organization"
                            value={companyName}
                            onChange={(e) => setCompanyName(e.target.value)}
                            className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                          />
                        </div>

                        <div className="space-y-2">
                          <Label htmlFor="industry" className="text-slate-700 font-medium">
                            Industry <span className="text-slate-400 font-normal">(optional)</span>
                          </Label>
                          <Input
                            id="industry"
                            placeholder="e.g., Technology, Marketing, Healthcare"
                            value={industry}
                            onChange={(e) => setIndustry(e.target.value)}
                            className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                          />
                        </div>
                      </>
                    )}

                    {/* Freelancer-specific fields */}
                    {role === 'FREELANCER' && (
                      <>
                        {/* Category Selection */}
                        <div className="space-y-2">
                          <Label className="text-slate-700 font-medium">Primary Category</Label>
                          <div className="grid grid-cols-2 gap-3">
                            {FREELANCER_CATEGORIES.map((cat) => (
                              <button
                                key={cat.value}
                                type="button"
                                onClick={() => setCategory(cat.value)}
                                className={`p-3 rounded-xl border-2 text-left transition-all duration-200 ${category === cat.value
                                  ? 'border-emerald-500 bg-emerald-50'
                                  : 'border-slate-200 hover:border-emerald-300'
                                  }`}
                              >
                                <div className="flex items-center gap-2.5">
                                  <cat.icon className={`w-5 h-5 ${category === cat.value ? cat.color : 'text-slate-400'}`} />
                                  <span className={`text-sm font-medium ${category === cat.value ? 'text-emerald-700' : 'text-slate-700'}`}>
                                    {cat.label}
                                  </span>
                                </div>
                              </button>
                            ))}
                          </div>
                        </div>

                        {/* Skills */}
                        <div className="space-y-2">
                          <Label htmlFor="skills" className="text-slate-700 font-medium">
                            Skills <span className="text-slate-400 font-normal">(up to 10)</span>
                          </Label>
                          <div className="flex gap-2">
                            <Input
                              id="skills"
                              placeholder="Type a skill and press Enter"
                              value={skillInput}
                              onChange={(e) => setSkillInput(e.target.value)}
                              onKeyDown={handleSkillKeyDown}
                              className="h-10 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                            />
                            <Button
                              type="button"
                              variant="outline"
                              onClick={addSkill}
                              className="shrink-0 border-slate-300 hover:border-emerald-300 hover:bg-emerald-50"
                              disabled={!skillInput.trim()}
                            >
                              Add
                            </Button>
                          </div>
                          {skills.length > 0 && (
                            <div className="flex flex-wrap gap-2 mt-2">
                              {skills.map((skill) => (
                                <Badge
                                  key={skill}
                                  variant="secondary"
                                  className="bg-emerald-100 text-emerald-700 border-emerald-200 pl-2.5 pr-1 py-1.5 gap-1"
                                >
                                  {skill}
                                  <button
                                    onClick={() => removeSkill(skill)}
                                    className="ml-0.5 w-4 h-4 rounded-full hover:bg-emerald-200 flex items-center justify-center transition-colors"
                                  >
                                    <span className="text-emerald-600 text-xs leading-none">&times;</span>
                                  </button>
                                </Badge>
                              ))}
                            </div>
                          )}
                        </div>

                        {/* Hourly Rate */}
                        <div className="space-y-2">
                          <Label htmlFor="hourly-rate" className="text-slate-700 font-medium">
                            Hourly Rate (USD) <span className="text-slate-400 font-normal">(optional)</span>
                          </Label>
                          <div className="relative">
                            <DollarSign className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                            <Input
                              id="hourly-rate"
                              type="number"
                              placeholder="25"
                              value={hourlyRate}
                              onChange={(e) => setHourlyRate(e.target.value)}
                              className="h-11 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20 pl-9"
                              min="1"
                              max="500"
                            />
                          </div>
                        </div>

                        {/* Student Status */}
                        <div className="space-y-3">
                          <div className="flex items-center gap-2">
                            <Checkbox
                              id="is-student"
                              checked={isStudent}
                              onCheckedChange={(checked) => setIsStudent(checked === true)}
                              className="border-slate-300 data-[state=checked]:bg-emerald-500 data-[state=checked]:border-emerald-500"
                            />
                            <Label htmlFor="is-student" className="text-slate-700 font-medium cursor-pointer flex items-center gap-2">
                              <School className="w-4 h-4 text-emerald-500" />
                              I am currently a student
                            </Label>
                          </div>

                          {isStudent && (
                            <motion.div
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                            >
                              <Input
                                id="institution"
                                placeholder="Your university or institution"
                                value={institution}
                                onChange={(e) => setInstitution(e.target.value)}
                                className="h-10 border-slate-300 focus:border-emerald-500 focus:ring-emerald-500/20"
                              />
                            </motion.div>
                          )}
                        </div>

                        {/* Bio */}
                        <div className="space-y-2">
                          <Label htmlFor="bio" className="text-slate-700 font-medium">
                            Short Bio <span className="text-slate-400 font-normal">(optional)</span>
                          </Label>
                          <textarea
                            id="bio"
                            placeholder="Tell clients a bit about yourself and what you do..."
                            value={bio}
                            onChange={(e) => setBio(e.target.value)}
                            rows={3}
                            maxLength={300}
                            className="flex w-full rounded-lg border border-slate-300 bg-transparent px-3 py-2.5 text-sm shadow-xs placeholder:text-slate-400 focus:border-emerald-500 focus:ring-emerald-500/20 focus-visible:outline-none resize-none transition-colors"
                          />
                          <p className="text-xs text-slate-400 text-right">{bio.length}/300</p>
                        </div>
                      </>
                    )}

                    {/* Terms */}
                    <div className="flex items-start gap-2.5 pt-2">
                      <Checkbox
                        id="terms"
                        checked={agreedToTerms}
                        onCheckedChange={(checked) => setAgreedToTerms(checked === true)}
                        className="border-slate-300 data-[state=checked]:bg-emerald-500 data-[state=checked]:border-emerald-500 mt-0.5"
                      />
                      <Label htmlFor="terms" className="text-sm text-slate-500 leading-relaxed cursor-pointer">
                        I agree to the{' '}
                        <button type="button" className="text-emerald-600 hover:text-emerald-700 font-medium underline underline-offset-2">
                          Terms of Service
                        </button>{' '}
                        and{' '}
                        <button type="button" className="text-emerald-600 hover:text-emerald-700 font-medium underline underline-offset-2">
                          Privacy Policy
                        </button>
                      </Label>
                    </div>
                  </CardContent>
                  <CardFooter className="flex flex-col gap-3">
                    <Button
                      onClick={handleSubmit}
                      disabled={isSubmitting}
                      className="w-full h-11 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg font-semibold shadow-lg shadow-emerald-500/25"
                    >
                      {isSubmitting ? (
                        <>
                          <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                          Creating account...
                        </>
                      ) : (
                        <>
                          Create Account
                          <ArrowRight className="w-4 h-4 ml-1" />
                        </>
                      )}
                    </Button>
                    <p className="text-sm text-slate-500 text-center">
                      Already have an account?{' '}
                      <button
                        onClick={() => navigate('auth/login')}
                        className="text-emerald-600 hover:text-emerald-700 font-semibold"
                      >
                        Sign in
                      </button>
                    </p>
                  </CardFooter>
                </Card>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
    </div>
  )
}
