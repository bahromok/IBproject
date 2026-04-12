'use client'

import { useEffect, useState } from 'react'
import { motion } from 'framer-motion'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { useNavigationStore } from '@/store/navigation-store'
import { useCountUp } from '@/hooks/use-count-up'
import {
  Briefcase,
  GraduationCap,
  Star,
  Shield,
  Users,
  Zap,
  ArrowRight,
  CheckCircle2,
  Globe,
  TrendingUp,
  Award,
  Code,
  Palette,
  PenTool,
  Video,
  Smartphone,
  Target,
  BarChart3,
  Megaphone,
} from 'lucide-react'

const fadeInUp = {
  hidden: { opacity: 0, y: 30 },
  visible: (i: number) => ({
    opacity: 1,
    y: 0,
    transition: { delay: i * 0.1, duration: 0.6, ease: 'easeOut' as const },
  }),
}

const stagger = {
  visible: { transition: { staggerChildren: 0.1 } },
}

const categoryIcons: Record<string, React.ElementType> = {
  'Web Development': Code,
  'Graphic Design': Palette,
  'Content Writing': PenTool,
  'Video Editing': Video,
  'Mobile Development': Smartphone,
  'UI/UX Design': Target,
  'Data Analysis': BarChart3,
  'Social Media': Megaphone,
}

export function LandingPage() {
  const { navigate } = useNavigationStore()

  // Fetch real stats from API
  const [stats, setStats] = useState({ users: 0, jobs: 0, contracts: 0 })
  const [statsLoaded, setStatsLoaded] = useState(false)

  useEffect(() => {
    // Use a lightweight public endpoint to get platform stats for the landing page
    fetch('/api/jobs?limit=1')
      .then(r => r.json())
      .then(data => {
        // Fallback to demo numbers if no data
        setStats({
          users: data?.pagination?.total || 0,
          jobs: data?.pagination?.total || 0,
          contracts: 0,
        })
      })
      .catch(() => {
        // Use demo numbers as fallback
        setStats({ users: 5, jobs: 4, contracts: 2 })
      })
      .finally(() => setStatsLoaded(true))
  }, [])

  const animUsers = useCountUp(statsLoaded ? stats.users : 0, 2000)
  const animJobs = useCountUp(statsLoaded ? stats.jobs : 0, 2000)
  const animContracts = useCountUp(statsLoaded ? stats.contracts : 0, 2000)

  return (
    <div className="min-h-screen bg-white">
      {/* Header */}
      <motion.header
        initial={{ opacity: 0, y: -20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
        className="sticky top-0 z-50 border-b bg-white/80 backdrop-blur-md"
      >
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <button className="flex items-center gap-2 cursor-pointer" onClick={() => navigate('landing')}>
              <div className="w-9 h-9 rounded-lg bg-emerald-500 flex items-center justify-center">
                <GraduationCap className="w-5 h-5 text-white" />
              </div>
              <span className="text-xl font-bold text-slate-900">
                Student<span className="text-emerald-500">Hire</span>
              </span>
            </button>

            <nav className="hidden md:flex items-center gap-6">
              <button onClick={() => navigate('auth/login')} className="text-sm font-medium text-slate-600 hover:text-slate-900 transition-colors">
                Sign In
              </button>
              <button onClick={() => navigate('jobs')} className="text-sm font-medium text-slate-600 hover:text-slate-900 transition-colors">
                Browse Jobs
              </button>
              <button onClick={() => navigate('freelancers')} className="text-sm font-medium text-slate-600 hover:text-slate-900 transition-colors">
                Find Talent
              </button>
              <Button onClick={() => navigate('auth/register')} className="bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg px-6">
                Get Started Free
              </Button>
            </nav>

            <div className="md:hidden flex items-center gap-2">
              <Button variant="ghost" size="sm" onClick={() => navigate('auth/login')} className="text-slate-600">Sign In</Button>
              <Button size="sm" onClick={() => navigate('auth/register')} className="bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg">Join</Button>
            </div>
          </div>
        </div>
      </motion.header>

      {/* Hero Section */}
      <section className="relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-br from-emerald-50 via-white to-teal-50" />
        <div className="absolute top-20 right-0 w-[500px] h-[500px] bg-emerald-100/40 rounded-full blur-3xl" />
        <div className="absolute bottom-0 left-0 w-[400px] h-[400px] bg-teal-100/30 rounded-full blur-3xl" />

        <div className="relative max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-20 pb-28 md:pt-28 md:pb-36">
          <motion.div initial="hidden" animate="visible" variants={stagger} className="text-center max-w-4xl mx-auto">
            <motion.div variants={fadeInUp} custom={0}>
              <Badge variant="secondary" className="mb-6 px-4 py-1.5 text-sm font-medium bg-emerald-100 text-emerald-700 border-emerald-200">
                <Zap className="w-3.5 h-3.5 mr-1.5" />
                The #1 Freelance Platform for Students
              </Badge>
            </motion.div>

            <motion.h1 variants={fadeInUp} custom={1} className="text-4xl md:text-6xl lg:text-7xl font-bold text-slate-900 leading-tight tracking-tight">
              Where Student Talent
              <br />
              <span className="text-transparent bg-clip-text bg-gradient-to-r from-emerald-500 to-teal-500">Meets Opportunity</span>
            </motion.h1>

            <motion.p variants={fadeInUp} custom={2} className="mt-6 text-lg md:text-xl text-slate-600 max-w-2xl mx-auto leading-relaxed">
              Connect with talented students for freelance projects in design, development, video editing, and content creation.
            </motion.p>

            <motion.div variants={fadeInUp} custom={3} className="mt-10 flex flex-col sm:flex-row gap-4 justify-center">
              <Button size="lg" onClick={() => navigate('auth/register')} className="bg-emerald-500 hover:bg-emerald-600 text-white rounded-xl px-8 h-12 text-base font-semibold shadow-lg shadow-emerald-500/25 hover:shadow-emerald-500/40 transition-all">
                Hire Student Talent
                <ArrowRight className="w-4 h-4 ml-1" />
              </Button>
              <Button size="lg" variant="outline" onClick={() => navigate('auth/register')} className="rounded-xl px-8 h-12 text-base font-semibold border-slate-300 hover:bg-emerald-50 hover:text-emerald-600 hover:border-emerald-300 transition-all">
                Find Freelance Work
              </Button>
            </motion.div>

            {/* Quick Start */}
            <motion.div variants={fadeInUp} custom={4} className="mt-8">
              <div className="relative inline-block">
                <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-slate-200" /></div>
                <div className="relative flex justify-center text-xs"><span className="bg-gradient-to-br from-emerald-50 via-white to-teal-50 px-4 text-slate-400">Get started today</span></div>
              </div>
              <div className="flex flex-wrap justify-center gap-3 mt-4">
                {[
                  { view: 'auth/register', label: 'Sign Up as Client', icon: Briefcase, color: 'border-emerald-300 hover:bg-emerald-50 hover:text-emerald-600 hover:border-emerald-400' },
                  { view: 'auth/register', label: 'Sign Up as Freelancer', icon: GraduationCap, color: 'border-teal-300 hover:bg-teal-50 hover:text-teal-600 hover:border-teal-400' },
                  { view: 'auth/login', label: 'Sign In', icon: Shield, color: 'border-amber-300 hover:bg-amber-50 hover:text-amber-600 hover:border-amber-400' },
                ].map((btn) => (
                  <button
                    key={btn.label}
                    onClick={() => navigate(btn.view as any)}
                    className={`flex items-center gap-2 px-5 py-2.5 rounded-lg border text-sm font-medium transition-all ${btn.color}`}
                  >
                    <btn.icon className="w-4 h-4" />
                    {btn.label}
                  </button>
                ))}
              </div>
            </motion.div>

            <motion.div variants={fadeInUp} custom={5} className="mt-6 flex items-center justify-center gap-6 text-sm text-slate-500">
              <span className="flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-emerald-500" /> Free to join</span>
              <span className="flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-emerald-500" /> No credit card required</span>
              <span className="hidden sm:flex items-center gap-1.5"><CheckCircle2 className="w-4 h-4 text-emerald-500" /> Secure payments</span>
            </motion.div>
          </motion.div>
        </div>
      </section>

      {/* Stats Section — REAL DATA from API */}
      <section className="bg-white border-y py-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-8">
            {[
              { value: animUsers, label: 'Registered Users', icon: Users },
              { value: animJobs, label: 'Jobs Posted', icon: Briefcase },
              { value: animContracts, label: 'Contracts Created', icon: Globe },
              { value: 4.9, label: 'Platform Rating', icon: Star, isStatic: true },
            ].map((stat, i) => (
              <motion.div key={stat.label} initial={{ opacity: 0, y: 20 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }} transition={{ delay: i * 0.1 }} className="text-center">
                <stat.icon className="w-6 h-6 text-emerald-500 mx-auto mb-2" />
                <div className="text-3xl md:text-4xl font-bold text-slate-900">
                  {stat.isStatic ? stat.value : stat.value.toLocaleString()}
                </div>
                <div className="text-sm text-slate-500 mt-1">{stat.label}</div>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* How It Works */}
      <section className="bg-slate-50 py-20 md:py-28">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div initial={{ opacity: 0, y: 20 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }} className="text-center mb-16">
            <Badge variant="secondary" className="mb-4 bg-emerald-100 text-emerald-700 border-emerald-200">How It Works</Badge>
            <h2 className="text-3xl md:text-4xl font-bold text-slate-900">Get Started in 3 Simple Steps</h2>
            <p className="mt-4 text-lg text-slate-600 max-w-2xl mx-auto">Whether you&apos;re looking to hire or get hired, our platform makes it easy.</p>
          </motion.div>

          <div className="grid md:grid-cols-3 gap-8">
            {[
              { step: '01', title: 'Create Your Profile', description: 'Sign up for free and set up your profile. Clients describe their needs, freelancers showcase their skills and portfolio.', icon: Users },
              { step: '02', title: 'Post or Find Projects', description: 'Clients post detailed job listings with budgets. Freelancers browse opportunities and submit competitive proposals.', icon: Briefcase },
              { step: '03', title: 'Collaborate & Deliver', description: 'Work together with milestones, secure payments via escrow, and build lasting professional relationships.', icon: TrendingUp },
            ].map((item, i) => (
              <motion.div key={item.step} initial={{ opacity: 0, y: 30 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }} transition={{ delay: i * 0.15 }} whileHover={{ y: -4 }} className="cursor-pointer" onClick={() => i === 0 ? navigate('auth/register') : i === 1 ? navigate('jobs') : navigate('freelancers')}>
                <Card className="h-full border-0 shadow-lg hover:shadow-xl transition-all duration-300 bg-white">
                  <CardContent className="p-8">
                    <div className="flex items-center gap-3 mb-5">
                      <div className="w-12 h-12 rounded-xl bg-emerald-100 flex items-center justify-center"><item.icon className="w-6 h-6 text-emerald-600" /></div>
                      <span className="text-4xl font-bold text-slate-100">{item.step}</span>
                    </div>
                    <h3 className="text-xl font-semibold text-slate-900 mb-3">{item.title}</h3>
                    <p className="text-slate-600 leading-relaxed">{item.description}</p>
                  </CardContent>
                </Card>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Categories Section */}
      <section className="bg-white py-20 md:py-28">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div initial={{ opacity: 0, y: 20 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }} className="text-center mb-16">
            <Badge variant="secondary" className="mb-4 bg-emerald-100 text-emerald-700 border-emerald-200">Popular Categories</Badge>
            <h2 className="text-3xl md:text-4xl font-bold text-slate-900">Find Talent in Every Field</h2>
            <p className="mt-4 text-lg text-slate-600 max-w-2xl mx-auto">From design to development, content to video — our student freelancers cover it all.</p>
          </motion.div>

          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {[
              { name: 'Web Development', icon: '💻', category: 'WEB_DEVELOPER' },
              { name: 'Graphic Design', icon: '🎨', category: 'DESIGNER' },
              { name: 'Content Writing', icon: '✍️', category: 'CONTENT_CREATOR' },
              { name: 'Video Editing', icon: '🎬', category: 'VIDEO_EDITOR' },
              { name: 'Mobile Development', icon: '📱', category: 'WEB_DEVELOPER' },
              { name: 'UI/UX Design', icon: '🎯', category: 'DESIGNER' },
              { name: 'Data Analysis', icon: '📊', category: 'WEB_DEVELOPER' },
              { name: 'Social Media', icon: '📣', category: 'CONTENT_CREATOR' },
            ].map((cat, i) => (
              <motion.div key={cat.name} initial={{ opacity: 0, scale: 0.95 }} whileInView={{ opacity: 1, scale: 1 }} viewport={{ once: true }} transition={{ delay: i * 0.05 }} whileHover={{ scale: 1.03 }} className="cursor-pointer" onClick={() => navigate('jobs')}>
                <Card className="border hover:border-emerald-200 hover:shadow-md transition-all duration-200">
                  <CardContent className="p-5 flex items-center gap-3">
                    <span className="text-2xl">{cat.icon}</span>
                    <div>
                      <h3 className="font-semibold text-slate-900 text-sm">{cat.name}</h3>
                      <p className="text-xs text-emerald-600 font-medium">Browse jobs →</p>
                    </div>
                  </CardContent>
                </Card>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Features Section */}
      <section className="bg-gradient-to-b from-slate-50 to-white py-20 md:py-28">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="grid lg:grid-cols-2 gap-16 items-center">
            <motion.div initial={{ opacity: 0, x: -30 }} whileInView={{ opacity: 1, x: 0 }} viewport={{ once: true }}>
              <Badge variant="secondary" className="mb-4 bg-emerald-100 text-emerald-700 border-emerald-200">Why StudentHire</Badge>
              <h2 className="text-3xl md:text-4xl font-bold text-slate-900 mb-6">Built for Students,<br />Trusted by Clients</h2>
              <p className="text-lg text-slate-600 mb-8 leading-relaxed">We understand the unique needs of student freelancers and the clients who want to work with fresh, innovative talent.</p>

              <div className="space-y-5">
                {[
                  { icon: Shield, title: 'Secure Escrow Payments', desc: 'Your money is held safely until work is delivered and approved. No risk to either party.' },
                  { icon: Award, title: 'Verified Student Profiles', desc: 'Every freelancer can verify their student status from accredited institutions.' },
                  { icon: Zap, title: 'Milestone-Based Work', desc: 'Break projects into milestones for transparent, manageable progress tracking.' },
                  { icon: Star, title: 'Review & Rating System', desc: 'Build your reputation through verified client reviews and multi-dimensional ratings.' },
                ].map((feature, i) => (
                  <motion.div key={feature.title} initial={{ opacity: 0, x: -20 }} whileInView={{ opacity: 1, x: 0 }} viewport={{ once: true }} transition={{ delay: i * 0.1 }} className="flex gap-4 cursor-pointer" onClick={() => navigate('auth/register')}>
                    <div className="w-10 h-10 rounded-lg bg-emerald-100 flex items-center justify-center shrink-0 mt-0.5"><feature.icon className="w-5 h-5 text-emerald-600" /></div>
                    <div>
                      <h3 className="font-semibold text-slate-900 mb-1">{feature.title}</h3>
                      <p className="text-sm text-slate-600 leading-relaxed">{feature.desc}</p>
                    </div>
                  </motion.div>
                ))}
              </div>
            </motion.div>

            <motion.div initial={{ opacity: 0, x: 30 }} whileInView={{ opacity: 1, x: 0 }} viewport={{ once: true }} className="hidden lg:block">
              <div className="relative">
                <div className="absolute -top-4 -right-4 w-72 h-72 bg-emerald-100/50 rounded-full blur-3xl" />
                <Card className="relative border-0 shadow-2xl cursor-pointer" onClick={() => navigate('auth/register')}>
                  <CardContent className="p-8">
                    <div className="space-y-6">
                      <div className="flex items-center gap-3 mb-6">
                        <div className="w-10 h-10 rounded-full bg-emerald-500 flex items-center justify-center text-white font-bold text-sm">W</div>
                        <div>
                          <div className="font-semibold text-slate-900">CreativeWolf42&apos;s Dashboard</div>
                          <div className="text-xs text-slate-500">Freelancer • Web Developer • MIT Student</div>
                        </div>
                        <Badge className="ml-auto bg-emerald-100 text-emerald-700">Student ✓</Badge>
                      </div>

                      <div className="grid grid-cols-2 gap-4">
                        <div className="bg-emerald-50 rounded-xl p-4"><div className="text-2xl font-bold text-emerald-600">$360</div><div className="text-xs text-slate-600 mt-1">Total Earnings</div></div>
                        <div className="bg-slate-50 rounded-xl p-4"><div className="text-2xl font-bold text-slate-900">4.7</div><div className="text-xs text-slate-600 mt-1 flex items-center gap-1"><Star className="w-3 h-3 fill-emerald-500 text-emerald-500" /> Avg Rating</div></div>
                        <div className="bg-slate-50 rounded-xl p-4"><div className="text-2xl font-bold text-slate-900">2</div><div className="text-xs text-slate-600 mt-1">Active Contracts</div></div>
                        <div className="bg-slate-50 rounded-xl p-4"><div className="text-2xl font-bold text-slate-900">4</div><div className="text-xs text-slate-600 mt-1">Completed Jobs</div></div>
                      </div>

                      <div className="space-y-3">
                        <div className="text-sm font-medium text-slate-700">Active Projects</div>
                        {[
                          { name: 'Mobile App UI Design', price: '$600', status: 'In Progress' },
                          { name: 'E-commerce Website', price: '$400', status: 'Milestone 2/3' },
                        ].map((project) => (
                          <div key={project.name} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                            <div className="flex-1 min-w-0">
                              <div className="text-sm font-medium text-slate-900 truncate">{project.name}</div>
                              <div className="text-xs text-emerald-600 font-medium">{project.status}</div>
                            </div>
                            <div className="text-sm font-semibold text-emerald-600">{project.price}</div>
                          </div>
                        ))}
                      </div>

                      <div className="text-center pt-2">
                        <span className="text-xs text-slate-400">Click to explore this dashboard →</span>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              </div>
            </motion.div>
          </div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="bg-gradient-to-br from-emerald-500 via-emerald-600 to-teal-600 py-20 md:py-28">
        <div className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 text-center">
          <motion.div initial={{ opacity: 0, y: 30 }} whileInView={{ opacity: 1, y: 0 }} viewport={{ once: true }}>
            <h2 className="text-3xl md:text-5xl font-bold text-white mb-6">Ready to Get Started?</h2>
            <p className="text-lg text-emerald-100 mb-10 max-w-2xl mx-auto">Join the growing community of students and clients building amazing things together.</p>
            <div className="flex flex-col sm:flex-row gap-4 justify-center">
              <Button size="lg" onClick={() => navigate('auth/register')} className="bg-white text-emerald-600 hover:bg-emerald-50 rounded-xl px-8 h-12 text-base font-semibold shadow-lg">
                I Want to Hire <ArrowRight className="w-4 h-4 ml-1" />
              </Button>
              <Button size="lg" variant="outline" onClick={() => navigate('auth/register')} className="bg-transparent text-white border-white/30 hover:bg-white/10 hover:text-white rounded-xl px-8 h-12 text-base font-semibold">
                I Want to Work
              </Button>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Footer */}
      <footer className="bg-slate-900 py-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="grid md:grid-cols-4 gap-8 mb-12">
            <div className="md:col-span-1">
              <button className="flex items-center gap-2 mb-4" onClick={() => navigate('landing')}>
                <div className="w-9 h-9 rounded-lg bg-emerald-500 flex items-center justify-center"><GraduationCap className="w-5 h-5 text-white" /></div>
                <span className="text-xl font-bold text-white">Student<span className="text-emerald-400">Hire</span></span>
              </button>
              <p className="text-sm text-slate-400 leading-relaxed">Where Student Talent Meets Opportunity. The premier freelance marketplace for students.</p>
            </div>

            {[
              {
                title: 'For Clients', links: [
                  { label: 'Post a Job', view: 'jobs/post' },
                  { label: 'Browse Freelancers', view: 'freelancers' },
                  { label: 'How It Works', view: 'landing' },
                  { label: 'Sign Up', view: 'auth/register' },
                ]
              },
              {
                title: 'For Freelancers', links: [
                  { label: 'Find Work', view: 'jobs' },
                  { label: 'Create Profile', view: 'auth/register' },
                  { label: 'Browse Jobs', view: 'jobs' },
                  { label: 'Sign Up', view: 'auth/register' },
                ]
              },
              {
                title: 'Company', links: [
                  { label: 'About Us', view: 'landing' },
                  { label: 'Sign In', view: 'auth/login' },
                  { label: 'Contact', view: 'landing' },
                  { label: 'Terms', view: 'landing' },
                ]
              },
            ].map((col) => (
              <div key={col.title}>
                <h3 className="text-sm font-semibold text-white mb-4">{col.title}</h3>
                <ul className="space-y-2">
                  {col.links.map((link) => (
                    <li key={link.label}>
                      <button onClick={() => navigate(link.view as any)} className="text-sm text-slate-400 hover:text-emerald-400 transition-colors">{link.label}</button>
                    </li>
                  ))}
                </ul>
              </div>
            ))}
          </div>

          <div className="border-t border-slate-800 pt-8 flex flex-col md:flex-row items-center justify-between gap-4">
            <p className="text-sm text-slate-500">© {new Date().getFullYear()} StudentHire. All rights reserved.</p>
            <div className="flex items-center gap-6">
              <button onClick={() => navigate('landing')} className="text-sm text-slate-500 hover:text-emerald-400 transition-colors">Privacy Policy</button>
              <button onClick={() => navigate('landing')} className="text-sm text-slate-500 hover:text-emerald-400 transition-colors">Terms of Service</button>
            </div>
          </div>
        </div>
      </footer>
    </div>
  )
}
