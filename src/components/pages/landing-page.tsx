'use client'

import { useEffect, useState } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { Badge } from '@/components/ui/badge'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import {
  Code, Palette, Video, PenTool, Layout, Share2,
  Star, ArrowRight, Briefcase, Users, Trophy,
  CheckCircle, MessageSquare, Shield,
  Zap, FileText, TrendingUp
} from 'lucide-react'

const fadeInUp = {
  initial: { opacity: 0, y: 30 },
  animate: { opacity: 1, y: 0 },
  transition: { duration: 0.6 },
}

const stagger = {
  animate: { transition: { staggerChildren: 0.1 } },
}

const categories = [
  { icon: Code, label: 'Web Development', desc: 'Build websites, apps, and more', color: 'text-emerald-600', bg: 'bg-emerald-50' },
  { icon: Palette, label: 'Graphic Design', desc: 'Logos, branding, illustrations', color: 'text-pink-600', bg: 'bg-pink-50' },
  { icon: Video, label: 'Video Editing', desc: 'YouTube, social media, commercials', color: 'text-purple-600', bg: 'bg-purple-50' },
  { icon: PenTool, label: 'Content Writing', desc: 'Blogs, copywriting, social posts', color: 'text-amber-600', bg: 'bg-amber-50' },
  { icon: Layout, label: 'UI/UX Design', desc: 'User interfaces, wireframes, prototyping', color: 'text-cyan-600', bg: 'bg-cyan-50' },
  { icon: Share2, label: 'Social Media', desc: 'Management, strategy, content creation', color: 'text-rose-600', bg: 'bg-rose-50' },
]

const steps = [
  {
    icon: FileText,
    title: 'Post a Job or Create a Profile',
    desc: 'Clients post detailed job listings while talented students create comprehensive profiles showcasing their skills and experience.',
  },
  {
    icon: Users,
    title: 'Match & Connect',
    desc: 'Browse, filter, and find the perfect match using our smart search tools. Connect directly with candidates or clients.',
  },
  {
    icon: TrendingUp,
    title: 'Collaborate & Deliver',
    desc: 'Use milestones, real-time messaging, and secure payments to ensure smooth collaboration from start to finish.',
  },
]

const testimonials = [
  {
    name: 'Sarah Chen',
    role: 'Startup Founder',
    text: 'StudentHire connected us with incredible design talent. Our entire brand identity was created by a talented college student at a fraction of the cost. The quality exceeded our expectations.',
    rating: 5,
    avatar: 'SC',
  },
  {
    name: 'Marcus Johnson',
    role: 'Web Developer',
    text: 'As a computer science student, StudentHire gave me real-world experience while earning income. I\'ve built 15+ projects and now have a portfolio that helped me land a full-time position.',
    rating: 5,
    avatar: 'MJ',
  },
  {
    name: 'Emily Rodriguez',
    role: 'Marketing Director',
    text: 'We consistently find creative, hungry talent on StudentHire. The students bring fresh perspectives and up-to-date skills. It\'s been a game-changer for our content strategy.',
    rating: 5,
    avatar: 'ER',
  },
]

interface FeaturedFreelancer {
  id: string
  displayName: string
  title: string | null
  category: string
  rating: number
  hourlyRate: number | null
  avatar: string | null
  skills: string[]
}

export default function LandingPage() {
  const { navigate } = useNavigationStore()
  const [featured, setFeatured] = useState<FeaturedFreelancer[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    async function fetchFeatured() {
      try {
        const res = await fetch('/api/freelancers?isFeatured=true&limit=6&sortBy=rating')
        const data = await res.json()
        if (!cancelled) {
          if (data.freelancers) {
            setFeatured(data.freelancers.slice(0, 6))
          }
        }
      } catch {
        if (!cancelled) {
          setFeatured([
            { id: '1', displayName: 'CreativeFox42', title: 'Full-Stack Developer', category: 'WEB_DEVELOPER', rating: 4.9, hourlyRate: 45, avatar: null, skills: ['React', 'Node.js', 'TypeScript'] },
            { id: '2', displayName: 'PixelPanda88', title: 'UI/UX Designer', category: 'DESIGNER', rating: 4.8, hourlyRate: 35, avatar: null, skills: ['Figma', 'Adobe XD', 'Prototyping'] },
            { id: '3', displayName: 'VideoViper21', title: 'Video Editor & Motion Designer', category: 'VIDEO_EDITOR', rating: 4.7, hourlyRate: 30, avatar: null, skills: ['Premiere Pro', 'After Effects', 'DaVinci'] },
            { id: '4', displayName: 'WordWizard55', title: 'Content Strategist & Writer', category: 'CONTENT_CREATOR', rating: 4.9, hourlyRate: 25, avatar: null, skills: ['SEO', 'Copywriting', 'Blogging'] },
            { id: '5', displayName: 'DesignDove77', title: 'Graphic Designer & Illustrator', category: 'DESIGNER', rating: 4.8, hourlyRate: 40, avatar: null, skills: ['Photoshop', 'Illustrator', 'Branding'] },
            { id: '6', displayName: 'CodeNinja33', title: 'Mobile App Developer', category: 'WEB_DEVELOPER', rating: 4.6, hourlyRate: 50, avatar: null, skills: ['React Native', 'Flutter', 'Swift'] },
          ])
        }
      } finally {
        if (!cancelled) setLoading(false)
      }
    }
    fetchFeatured()
    return () => { cancelled = true }
  }, [])

  return (
    <div className="min-h-screen">
      {/* Navigation Bar */}
      <motion.nav
        initial={{ opacity: 0, y: -20 }}
        animate={{ opacity: 1, y: 0 }}
        className="fixed top-0 left-0 right-0 z-50 bg-white/80 backdrop-blur-md border-b border-border"
      >
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div
              className="flex items-center gap-2 cursor-pointer"
              onClick={() => navigate('landing')}
            >
              <img src="/logo.svg" alt="StudentHire" className="w-8 h-8 rounded-lg" />
              <span className="text-xl font-bold text-emerald-600">StudentHire</span>
            </div>
            <div className="hidden md:flex items-center gap-8">
              <button onClick={() => navigate('jobs')} className="text-sm font-medium text-muted-foreground hover:text-foreground transition-colors">
                Find Work
              </button>
              <button onClick={() => navigate('freelancers')} className="text-sm font-medium text-muted-foreground hover:text-foreground transition-colors">
                Find Talent
              </button>
              <button className="text-sm font-medium text-muted-foreground hover:text-foreground transition-colors">
                How it Works
              </button>
            </div>
            <div className="flex items-center gap-3">
              <Button variant="ghost" onClick={() => navigate('auth/login')}>
                Log In
              </Button>
              <Button className="bg-emerald-600 hover:bg-emerald-700 text-white" onClick={() => navigate('auth/register')}>
                Sign Up
              </Button>
            </div>
          </div>
        </div>
      </motion.nav>

      {/* Hero Section */}
      <section className="hero-gradient pt-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-20 md:py-32">
          <motion.div
            initial={{ opacity: 0, y: 40 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
            className="text-center max-w-4xl mx-auto"
          >
            <motion.div
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              transition={{ duration: 0.5, delay: 0.2 }}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm font-medium mb-8"
            >
              <Zap className="w-4 h-4" />
              Trusted by 10,000+ students and growing
            </motion.div>

            <h1 className="text-4xl sm:text-5xl md:text-6xl lg:text-7xl font-bold tracking-tight text-gray-900 leading-[1.1] mb-6">
              Where Student Talent
              <br />
              <span className="text-emerald-600">Meets Opportunity</span>
            </h1>

            <p className="text-lg sm:text-xl text-muted-foreground max-w-2xl mx-auto mb-10 leading-relaxed">
              Hire talented student designers, developers, editors, and creators — or find your next gig as a student freelancer.
            </p>

            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.6, delay: 0.4 }}
              className="flex flex-col sm:flex-row items-center justify-center gap-4"
            >
              <Button
                size="lg"
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-8 py-6 text-lg rounded-xl shadow-lg shadow-emerald-600/25 transition-all hover:shadow-xl hover:shadow-emerald-600/30 hover:-translate-y-0.5"
                onClick={() => navigate('freelancers')}
              >
                Find Talent
                <ArrowRight className="w-5 h-5 ml-2" />
              </Button>
              <Button
                size="lg"
                variant="outline"
                className="px-8 py-6 text-lg rounded-xl border-2 hover:bg-emerald-50 hover:border-emerald-300 transition-all hover:-translate-y-0.5"
                onClick={() => navigate('jobs')}
              >
                Find Work
                <Briefcase className="w-5 h-5 ml-2" />
              </Button>
            </motion.div>
          </motion.div>

          {/* Stats Bar */}
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.6 }}
            className="mt-16 md:mt-24 grid grid-cols-3 gap-4 md:gap-8 max-w-3xl mx-auto"
          >
            {[
              { label: 'Students', value: '10,000+', icon: Users },
              { label: 'Projects', value: '5,000+', icon: Briefcase },
              { label: 'Satisfaction', value: '98%', icon: Trophy },
            ].map((stat) => (
              <div key={stat.label} className="text-center p-4 rounded-xl bg-white/60 backdrop-blur-sm border border-white/80">
                <stat.icon className="w-5 h-5 text-emerald-600 mx-auto mb-2" />
                <div className="text-2xl md:text-3xl font-bold text-gray-900">{stat.value}</div>
                <div className="text-sm text-muted-foreground mt-1">{stat.label}</div>
              </div>
            ))}
          </motion.div>
        </div>
      </section>

      {/* How It Works */}
      <section className="py-20 md:py-28 bg-white">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="text-center mb-16"
          >
            <motion.p variants={fadeInUp} className="text-sm font-semibold text-emerald-600 uppercase tracking-wider mb-3">
              How it works
            </motion.p>
            <motion.h2 variants={fadeInUp} className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
              Three Simple Steps
            </motion.h2>
            <motion.p variants={fadeInUp} className="text-lg text-muted-foreground max-w-2xl mx-auto">
              Get started in minutes. Whether you&apos;re hiring or looking for work, our platform makes it easy.
            </motion.p>
          </motion.div>

          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="grid md:grid-cols-3 gap-8"
          >
            {steps.map((step, i) => (
              <motion.div key={step.title} variants={fadeInUp}>
                <Card className="relative border-0 shadow-lg hover:shadow-xl transition-shadow h-full">
                  <CardContent className="p-8">
                    <div className="flex items-center gap-4 mb-6">
                      <div className="w-12 h-12 rounded-xl bg-emerald-50 flex items-center justify-center">
                        <step.icon className="w-6 h-6 text-emerald-600" />
                      </div>
                      <span className="text-sm font-bold text-emerald-600 bg-emerald-50 px-3 py-1 rounded-full">
                        Step {i + 1}
                      </span>
                    </div>
                    <h3 className="text-xl font-semibold text-gray-900 mb-3">{step.title}</h3>
                    <p className="text-muted-foreground leading-relaxed">{step.desc}</p>
                  </CardContent>
                </Card>
              </motion.div>
            ))}
          </motion.div>
        </div>
      </section>

      {/* Categories Section */}
      <section className="py-20 md:py-28 bg-gray-50/50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="text-center mb-16"
          >
            <motion.p variants={fadeInUp} className="text-sm font-semibold text-emerald-600 uppercase tracking-wider mb-3">
              Categories
            </motion.p>
            <motion.h2 variants={fadeInUp} className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
              Browse by Category
            </motion.h2>
            <motion.p variants={fadeInUp} className="text-lg text-muted-foreground max-w-2xl mx-auto">
              Find the perfect freelancer for your project across a wide range of categories.
            </motion.p>
          </motion.div>

          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="grid grid-cols-2 md:grid-cols-3 gap-4 md:gap-6"
          >
            {categories.map((cat) => (
              <motion.div key={cat.label} variants={fadeInUp}>
                <Card
                  className="group cursor-pointer border-0 shadow-md hover:shadow-lg transition-all hover:-translate-y-1 h-full"
                  onClick={() => navigate('jobs')}
                >
                  <CardContent className="p-6 flex flex-col items-center text-center">
                    <div className={`w-14 h-14 rounded-2xl ${cat.bg} flex items-center justify-center mb-4 group-hover:scale-110 transition-transform`}>
                      <cat.icon className={`w-7 h-7 ${cat.color}`} />
                    </div>
                    <h3 className="font-semibold text-gray-900 mb-1">{cat.label}</h3>
                    <p className="text-sm text-muted-foreground">{cat.desc}</p>
                  </CardContent>
                </Card>
              </motion.div>
            ))}
          </motion.div>
        </div>
      </section>

      {/* Featured Freelancers */}
      <section className="py-20 md:py-28 bg-white">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="text-center mb-16"
          >
            <motion.p variants={fadeInUp} className="text-sm font-semibold text-emerald-600 uppercase tracking-wider mb-3">
              Top Talent
            </motion.p>
            <motion.h2 variants={fadeInUp} className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
              Featured Freelancers
            </motion.h2>
            <motion.p variants={fadeInUp} className="text-lg text-muted-foreground max-w-2xl mx-auto">
              Meet some of our top-rated student freelancers ready to bring your ideas to life.
            </motion.p>
          </motion.div>

          {loading ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
              {[1, 2, 3, 4, 5, 6].map((i) => (
                <Card key={i} className="border-0 shadow-md">
                  <CardContent className="p-6">
                    <div className="flex items-center gap-4 mb-4">
                      <div className="w-14 h-14 rounded-full bg-gray-100 animate-pulse" />
                      <div className="space-y-2 flex-1">
                        <div className="h-4 bg-gray-100 rounded animate-pulse w-3/4" />
                        <div className="h-3 bg-gray-100 rounded animate-pulse w-1/2" />
                      </div>
                    </div>
                    <div className="space-y-2">
                      <div className="h-3 bg-gray-100 rounded animate-pulse" />
                      <div className="h-3 bg-gray-100 rounded animate-pulse w-2/3" />
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          ) : (
            <motion.div
              initial="initial"
              whileInView="animate"
              viewport={{ once: true, margin: '-100px' }}
              variants={stagger}
              className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6"
            >
              {featured.map((f) => (
                <motion.div key={f.id} variants={fadeInUp}>
                  <Card className="border-0 shadow-md hover:shadow-lg transition-all hover:-translate-y-1 h-full">
                    <CardContent className="p-6">
                      <div className="flex items-start gap-4 mb-4">
                        <Avatar className="w-14 h-14 border-2 border-emerald-100">
                          <AvatarFallback className="bg-emerald-50 text-emerald-700 font-semibold text-sm">
                            {f.displayName.slice(0, 2).toUpperCase()}
                          </AvatarFallback>
                        </Avatar>
                        <div className="flex-1 min-w-0">
                          <h3 className="font-semibold text-gray-900 truncate">{f.displayName}</h3>
                          <p className="text-sm text-muted-foreground truncate">{f.title}</p>
                          <div className="flex items-center gap-1 mt-1">
                            <Star className="w-3.5 h-3.5 fill-amber-400 text-amber-400" />
                            <span className="text-sm font-medium">{f.rating.toFixed(1)}</span>
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center justify-between">
                        <Badge variant="secondary" className="text-xs">
                          {f.category.replace(/_/g, ' ')}
                        </Badge>
                        <span className="text-sm font-semibold text-emerald-600">
                          ${f.hourlyRate || 0}/hr
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-1 mt-3">
                        {(f.skills || []).slice(0, 3).map((skill) => (
                          <Badge key={skill} variant="outline" className="text-xs font-normal">
                            {skill}
                          </Badge>
                        ))}
                      </div>
                      <Button
                        variant="outline"
                        className="w-full mt-4 border-emerald-200 text-emerald-600 hover:bg-emerald-50 hover:text-emerald-700"
                        onClick={() => navigate('freelancers/detail', { id: f.id })}
                      >
                        View Profile
                      </Button>
                    </CardContent>
                  </Card>
                </motion.div>
              ))}
            </motion.div>
          )}

          <motion.div
            initial={{ opacity: 0 }}
            whileInView={{ opacity: 1 }}
            viewport={{ once: true }}
            className="text-center mt-12"
          >
            <Button
              size="lg"
              variant="outline"
              className="border-emerald-200 text-emerald-600 hover:bg-emerald-50 px-8"
              onClick={() => navigate('freelancers')}
            >
              Browse All Freelancers
              <ArrowRight className="w-4 h-4 ml-2" />
            </Button>
          </motion.div>
        </div>
      </section>

      {/* Testimonials */}
      <section className="py-20 md:py-28 bg-gray-50/50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="text-center mb-16"
          >
            <motion.p variants={fadeInUp} className="text-sm font-semibold text-emerald-600 uppercase tracking-wider mb-3">
              Testimonials
            </motion.p>
            <motion.h2 variants={fadeInUp} className="text-3xl md:text-4xl font-bold text-gray-900 mb-4">
              Loved by Clients & Freelancers
            </motion.h2>
          </motion.div>

          <motion.div
            initial="initial"
            whileInView="animate"
            viewport={{ once: true, margin: '-100px' }}
            variants={stagger}
            className="grid md:grid-cols-3 gap-6"
          >
            {testimonials.map((t) => (
              <motion.div key={t.name} variants={fadeInUp}>
                <Card className="border-0 shadow-md h-full">
                  <CardContent className="p-6">
                    <div className="flex gap-1 mb-4">
                      {Array.from({ length: t.rating }).map((_, i) => (
                        <Star key={i} className="w-4 h-4 fill-amber-400 text-amber-400" />
                      ))}
                    </div>
                    <p className="text-muted-foreground mb-6 leading-relaxed italic">
                      &ldquo;{t.text}&rdquo;
                    </p>
                    <div className="flex items-center gap-3">
                      <Avatar className="w-10 h-10">
                        <AvatarFallback className="bg-emerald-50 text-emerald-700 font-semibold text-xs">
                          {t.avatar}
                        </AvatarFallback>
                      </Avatar>
                      <div>
                        <div className="font-semibold text-sm text-gray-900">{t.name}</div>
                        <div className="text-xs text-muted-foreground">{t.role}</div>
                      </div>
                    </div>
                  </CardContent>
                </Card>
              </motion.div>
            ))}
          </motion.div>
        </div>
      </section>

      {/* CTA Section */}
      <section className="py-20 md:py-28 bg-white">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true, margin: '-100px' }}
            transition={{ duration: 0.6 }}
            className="relative overflow-hidden rounded-3xl bg-emerald-600 px-8 py-16 md:px-16 md:py-20 text-center"
          >
            <div className="absolute inset-0 bg-gradient-to-br from-emerald-600 via-emerald-700 to-teal-800" />
            <div className="absolute inset-0 opacity-10" style={{ backgroundImage: 'radial-gradient(circle at 25% 25%, white 1px, transparent 1px)', backgroundSize: '40px 40px' }} />
            <div className="relative z-10">
              <h2 className="text-3xl md:text-5xl font-bold text-white mb-6">
                Ready to get started?
              </h2>
              <p className="text-lg text-emerald-100 max-w-2xl mx-auto mb-10">
                Join thousands of students and clients already using StudentHire. Sign up today and start your journey.
              </p>
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <Button
                  size="lg"
                  className="bg-white text-emerald-700 hover:bg-emerald-50 px-8 py-6 text-lg rounded-xl shadow-lg transition-all hover:-translate-y-0.5"
                  onClick={() => navigate('auth/register')}
                >
                  Create Free Account
                  <ArrowRight className="w-5 h-5 ml-2" />
                </Button>
                <Button
                  size="lg"
                  variant="outline"
                  className="border-2 border-white/30 text-white hover:bg-white/10 px-8 py-6 text-lg rounded-xl transition-all hover:-translate-y-0.5"
                  onClick={() => navigate('jobs')}
                >
                  Browse Jobs
                </Button>
              </div>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Footer */}
      <footer className="bg-gray-900 text-gray-400 py-16">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-8 mb-12">
            <div className="col-span-2 md:col-span-1">
              <div className="flex items-center gap-2 mb-4">
                <img src="/logo.svg" alt="StudentHire" className="w-8 h-8 rounded-lg" />
                <span className="text-xl font-bold text-white">StudentHire</span>
              </div>
              <p className="text-sm leading-relaxed">
                The freelance marketplace built for students and the clients who need their talent.
              </p>
            </div>
            <div>
              <h4 className="font-semibold text-white mb-4 text-sm uppercase tracking-wider">Platform</h4>
              <ul className="space-y-3 text-sm">
                <li><button onClick={() => navigate('jobs')} className="hover:text-white transition-colors">Find Work</button></li>
                <li><button onClick={() => navigate('freelancers')} className="hover:text-white transition-colors">Find Talent</button></li>
                <li><button className="hover:text-white transition-colors">How it Works</button></li>
                <li><button className="hover:text-white transition-colors">Categories</button></li>
              </ul>
            </div>
            <div>
              <h4 className="font-semibold text-white mb-4 text-sm uppercase tracking-wider">Support</h4>
              <ul className="space-y-3 text-sm">
                <li><button className="hover:text-white transition-colors">Help Center</button></li>
                <li><button className="hover:text-white transition-colors">Contact Us</button></li>
                <li><button className="hover:text-white transition-colors">Safety</button></li>
                <li><button className="hover:text-white transition-colors">Community</button></li>
              </ul>
            </div>
            <div>
              <h4 className="font-semibold text-white mb-4 text-sm uppercase tracking-wider">Legal</h4>
              <ul className="space-y-3 text-sm">
                <li><button className="hover:text-white transition-colors">Terms of Service</button></li>
                <li><button className="hover:text-white transition-colors">Privacy Policy</button></li>
                <li><button className="hover:text-white transition-colors">Cookie Policy</button></li>
                <li><button className="hover:text-white transition-colors">Accessibility</button></li>
              </ul>
            </div>
          </div>
          <div className="border-t border-gray-800 pt-8 flex flex-col md:flex-row items-center justify-between gap-4">
            <p className="text-sm">&copy; {new Date().getFullYear()} StudentHire. All rights reserved.</p>
            <div className="flex items-center gap-6">
              <Shield className="w-4 h-4 text-emerald-500" />
              <span className="text-xs">Secure payments &middot; Identity protected &middot; 24/7 support</span>
            </div>
          </div>
        </div>
      </footer>
    </div>
  )
}
