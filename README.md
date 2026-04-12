<p align="center">
  <img src="public/logo.svg" alt="StudentHire Logo" width="200" height="200" />
</p>

<h1 align="center">StudentHire</h1>

<p align="center">
  <strong>Where Student Talent Meets Opportunity</strong>
</p>

<p align="center">
  A full-featured freelance marketplace platform built with Next.js, designed to connect student freelancers with clients who need their skills. Features identity protection, milestone-based payments, and a comprehensive admin dashboard.
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Next.js-16-black?logo=next.js" alt="Next.js 16" />
  <img src="https://img.shields.io/badge/TypeScript-5-blue?logo=typescript" alt="TypeScript 5" />
  <img src="https://img.shields.io/badge/Tailwind_CSS-4-06B6D4?logo=tailwindcss" alt="Tailwind CSS 4" />
  <img src="https://img.shields.io/badge/Prisma-6-2D3748?logo=prisma" alt="Prisma 6" />
  <img src="https://img.shields.io/badge/shadcn%2Fui-New%20York-18181B" alt="shadcn/ui" />
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License" />
</p>

---

## Table of Contents

- [Project Overview](#project-overview)
- [Key Features](#key-features)
- [Tech Stack](#tech-stack)
- [Quick Start (Local Development)](#quick-start-local-development-with-sqlite)
- [PostgreSQL Setup (Production)](#postgresql-setup-production)
- [Vercel Deployment Guide](#vercel-deployment-guide)
- [Project Structure](#project-structure)
- [Environment Variables](#environment-variables)
- [Demo Accounts](#demo-accounts)
- [Security Features](#security-features-fraud-prevention)
- [API Documentation](#api-documentation)

---

## Project Overview

**StudentHire** is a production-ready freelance marketplace built as a Single Page Application (SPA) using Next.js 15 App Router. The platform is specifically designed to:

- **Empower student freelancers** by connecting them with real-world projects and clients
- **Protect user identities** through a display name system that never exposes real names or emails between users
- **Provide safe payments** via an escrow system with milestone-based releases
- **Offer comprehensive moderation** through admin tools, report systems, and dispute resolution

The platform supports three user roles — **Client**, **Freelancer**, and **Admin** — each with dedicated dashboards, tools, and workflows.

## Key Features

### For Freelancers
- **Profile Builder** — Create a rich profile with portfolio, skills, student verification, and availability status
- **Job Marketplace** — Browse and search jobs with advanced filters (category, budget, experience level)
- **Proposal System** — Submit tailored proposals with cover letters and bid amounts
- **Contract Management** — Track milestones, submit work, and manage active contracts
- **In-Platform Messaging** — Communicate with clients through identity-protected conversations
- **Reviews & Ratings** — Build reputation through client reviews with detailed score breakdowns
- **Saved Jobs** — Bookmark interesting opportunities for later

### For Clients
- **Job Posting** — Multi-step job creation with skills, budget, duration, and experience level
- **Proposal Management** — Review, shortlist, accept, or reject freelancer proposals
- **Contract Dashboard** — Track active contracts, approve milestones, and release payments
- **Freelancer Discovery** — Browse freelancer profiles with filters for skills, ratings, and availability
- **Escrow Payments** — Deposit funds securely before work begins, release upon milestone approval
- **Review System** — Rate freelancers on quality, communication, and professionalism

### For Administrators
- **User Management** — Verify, suspend, and manage all platform users
- **Job Moderation** — Feature, review, and remove job listings
- **Report Handling** — Investigate and resolve user reports with admin notes
- **Dispute Resolution** — Mediate contract disputes with refund/release capabilities
- **Analytics Dashboard** — Full platform statistics (users, jobs, revenue, growth trends)

### Platform-Wide
- **Identity Protection** — Real emails and names are never exposed between users
- **Role-Based Access Control** — All API routes enforce proper authorization
- **Responsive Design** — Fully responsive with mobile sidebar navigation
- **Real-Time Notifications** — In-app notifications for proposals, messages, payments, and more
- **Fraud Prevention** — Escrow payments, report system, admin moderation, student verification

## Tech Stack

| Category | Technology | Purpose |
|---|---|---|
| **Framework** | Next.js 16 (App Router) | Full-stack React framework |
| **Language** | TypeScript 5 | Type-safe development |
| **Styling** | Tailwind CSS 4 | Utility-first CSS framework |
| **UI Components** | shadcn/ui (New York) | Pre-built accessible components |
| **Database ORM** | Prisma 6 | Type-safe database client |
| **Database** | SQLite (dev) / PostgreSQL (prod) | Data persistence |
| **Authentication** | NextAuth.js v4 | Session-based auth with JWT |
| **State Management** | Zustand | Client-side state |
| **Server State** | TanStack Query | Data fetching and caching |
| **Animations** | Framer Motion | Page transitions and UI animations |
| **Charts** | Recharts | Dashboard analytics charts |
| **Forms** | React Hook Form + Zod | Form handling and validation |
| **Icons** | Lucide React | Icon library |
| **Date Utilities** | date-fns | Date formatting and manipulation |
| **Password Hashing** | bcryptjs | Secure password storage |
| **Package Manager** | Bun | Fast JavaScript runtime and package manager |

---

## Quick Start (Local Development with SQLite)

### Prerequisites

- [Bun](https://bun.sh/) (recommended) or [Node.js](https://nodejs.org/) 18+
- Git

### Setup Steps

```bash
# 1. Clone the repository
git clone https://github.com/your-username/studenthire.git
cd studenthire

# 2. Install dependencies
bun install

# 3. Set up environment variables
cp .env.example .env
# Edit .env with the following:
#   DATABASE_URL="file:./dev.db"

# 4. Push the database schema
bun run db:push

# 5. Seed the database with demo data
bun run db:seed

# 6. Start the development server
bun run dev
```

The application will be available at [http://localhost:3000](http://localhost:3000).

### Available Scripts

| Command | Description |
|---|---|
| `bun run dev` | Start development server on port 3000 |
| `bun run build` | Build for production |
| `bun run start` | Start production server |
| `bun run lint` | Run ESLint |
| `bun run db:push` | Push Prisma schema to database |
| `bun run db:generate` | Generate Prisma Client |
| `bun run db:migrate` | Run database migrations |
| `bun run db:seed` | Seed the database with demo data |
| `bun run db:reset` | Reset the database |

---

## PostgreSQL Setup (Production)

### 3a. Install PostgreSQL

Choose the method for your operating system:

**macOS (Homebrew):**
```bash
brew install postgresql@16
brew services start postgresql@16
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt update
sudo apt install postgresql postgresql-contrib
sudo systemctl start postgresql
sudo systemctl enable postgresql
```

**Windows:**
1. Download the installer from [postgresql.org/download](https://www.postgresql.org/download/windows/)
2. Run the installer and follow the setup wizard
3. Set a password for the `postgres` superuser
4. Ensure the PostgreSQL service is running

### 3b. Create Database and User

```bash
# Access PostgreSQL as the superuser
sudo -u postgres psql
```

```sql
-- Create the database
CREATE DATABASE studenthire;

-- Create a dedicated user with a strong password
CREATE USER studenthire_user WITH PASSWORD 'your_secure_password';

-- Grant all privileges on the database
GRANT ALL PRIVILEGES ON DATABASE studenthire TO studenthire_user;

-- Connect to the new database and grant schema permissions
\c studenthire
GRANT ALL ON SCHEMA public TO studenthire_user;

-- Exit
\q
```

### 3c. Expose PostgreSQL for Development

By default, PostgreSQL only accepts local connections. To allow external connections (for tools like pgAdmin, DBeaver, or remote development):

#### Edit `postgresql.conf`

Find the configuration file for your installation:

| OS | Path |
|---|---|
| **macOS (Homebrew)** | `/opt/homebrew/var/postgresql@16/postgresql.conf` |
| **Linux (Ubuntu)** | `/etc/postgresql/16/main/postgresql.conf` |
| **Windows** | `C:\Program Files\PostgreSQL\16\data\postgresql.conf` |

Uncomment or add the following line:

```conf
# Accept connections from any IP (development only)
listen_addresses = '*'

# Confirm the port is set
port = 5432
```

#### Edit `pg_hba.conf`

The `pg_hba.conf` file is in the same directory as `postgresql.conf`. Add the following line to allow password-based connections:

```conf
# For development: accept connections from any IP
host    all             all             0.0.0.0/0               md5

# For better security: accept local connections only
host    all             all             127.0.0.1/32            md5
```

#### Restart PostgreSQL

```bash
# macOS
brew services restart postgresql@16

# Linux
sudo systemctl restart postgresql

# Windows (Command Prompt as Administrator)
net stop postgresql-x64-16 && net start postgresql-x64-16
```

### 3d. Test Connection

```bash
# Test with psql
psql -h localhost -U studenthire_user -d studenthire

# You should be prompted for the password and connected
```

If the connection fails, double-check the `pg_hba.conf` and `postgresql.conf` settings, and ensure the firewall allows traffic on port 5432.

### 3e. Configure Prisma for PostgreSQL

Update the datasource in `prisma/schema.prisma`:

```prisma
datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
}
```

Update your `.env` file:

```env
DATABASE_URL="postgresql://studenthire_user:your_secure_password@localhost:5432/studenthire"
```

### 3f. Run Migrations and Seed

```bash
# Push the schema to PostgreSQL
bun run db:push

# Seed with demo data
bun run db:seed
```

For production with migration history:

```bash
# Create and apply migration
bun run db:migrate

# Seed the database
bun run db:seed
```

### 3g. Connect from External Tools

Use these connection settings in database GUI tools (pgAdmin, DBeaver, TablePlus, etc.):

| Setting | Value |
|---|---|
| **Host** | `localhost` (or your machine's IP for remote access) |
| **Port** | `5432` |
| **Database** | `studenthire` |
| **Username** | `studenthire_user` |
| **Password** | `your_secure_password` |

### 3h. Expose PostgreSQL for Remote Connections (Optional)

> ⚠️ **Warning:** Exposing your database to the internet is not recommended for production. Use a managed database service (Neon, Supabase, Railway) for production deployments.

**Using ngrok (for temporary remote access):**
```bash
# Install ngrok: https://ngrok.com/download
ngrok tcp 5432
```

This will give you a public URL (e.g., `tcp://0.tcp.ngrok.io:12345`) that you can use as your host.

**For production, use:**
- **Connection pooling** with [PgBouncer](https://www.pgbouncer.org/) to manage database connections
- **Managed PostgreSQL** services like Neon, Supabase, or Railway
- **SSL/TLS encryption** (required by most managed services)

---

## Vercel Deployment Guide

### 4a. Prerequisites

- A [Vercel](https://vercel.com/) account (free tier works)
- A managed PostgreSQL database (recommended services below)
- A [GitHub](https://github.com/) repository with your code

### 4b. Set Up Remote Database

We recommend one of these managed PostgreSQL services:

| Service | Free Tier | Best For |
|---|---|---|
| [Neon](https://neon.tech) | 0.5 GB, 1 project | Serverless, branching |
| [Supabase](https://supabase.com) | 500 MB, 2 projects | Full backend platform |
| [Railway](https://railway.app) | $5 credit/month | Simple setup |

**Setup steps (using Neon as example):**

1. Go to [neon.tech](https://neon.tech) and sign up
2. Create a new project (choose the region closest to your users)
3. Copy the connection string — it will look like:
   ```
   postgresql://studenthire_user:your_password@ep-cool-name-12345.us-east-2.aws.neon.tech/studenthire?sslmode=require
   ```

### 4c. Push Schema to Remote Database

```bash
# Temporarily override DATABASE_URL to point to your remote database
DATABASE_URL="postgresql://user:pass@ep-xyz.region.aws.neon.tech/dbname?sslmode=require" bun run db:push
DATABASE_URL="postgresql://user:pass@ep-xyz.region.aws.neon.tech/dbname?sslmode=require" bun run db:seed
```

### 4d. Deploy via Vercel CLI

```bash
# Install Vercel CLI globally
npm i -g vercel

# Login to Vercel
vercel login

# Deploy (follow the prompts)
vercel

# Add environment variable
vercel env add DATABASE_URL
# Paste your remote PostgreSQL connection string when prompted

# Add other environment variables
vercel env add NEXTAUTH_URL
# Enter: https://your-project.vercel.app

vercel env add NEXTAUTH_SECRET
# Generate a secret: openssl rand -base64 32

# Deploy to production
vercel --prod
```

### 4e. Deploy via GitHub Integration

1. Push your code to GitHub
2. Go to [vercel.com](https://vercel.com) and sign in
3. Click **"New Project"**
4. Import your GitHub repository
5. Configure the project:
   - **Framework Preset:** Next.js
   - **Root Directory:** `.` (default)
   - **Build Command:** `npm run build` (auto-detected)
   - **Output Directory:** `.next` (auto-detected)
6. **Add Environment Variables:**

| Variable | Value |
|---|---|
| `DATABASE_URL` | Your PostgreSQL connection string (with `?sslmode=require`) |
| `NEXTAUTH_URL` | `https://your-domain.vercel.app` |
| `NEXTAUTH_SECRET` | Generate with `openssl rand -base64 32` |
| `NODE_ENV` | `production` |

7. Click **"Deploy"**

### 4f. Post-Deployment Checklist

- [ ] **Custom Domain** — Add your domain in Vercel project settings → Domains
- [ ] **Database SSL** — Ensure your `DATABASE_URL` includes `?sslmode=require`
- [ ] **CORS Configuration** — Update `NEXTAUTH_URL` if using a custom domain
- [ ] **Environment Variables** — Verify all env vars are set in Vercel dashboard
- [ ] **Database Backups** — Enable automatic backups on your managed database
- [ ] **Monitoring** — Enable Vercel Analytics and Speed Insights
- [ ] **Error Tracking** — Set up [Sentry](https://sentry.io/) or similar
- [ ] **Performance** — Review Lighthouse scores and optimize images

---

## Project Structure

```
studenthire/
├── prisma/
│   ├── schema.prisma          # Database schema (15 models, 18 enums)
│   └── seed.ts                # Comprehensive seed data script
├── public/
│   ├── favicon.ico            # App favicon
│   └── logo.svg               # StudentHire logo
├── src/
│   ├── app/
│   │   ├── api/               # All API routes (REST)
│   │   │   ├── auth/          # Authentication (register, NextAuth)
│   │   │   │   ├── register/  # POST /api/auth/register
│   │   │   │   └── [...nextauth]/ # GET/POST /api/auth/[...nextauth]
│   │   │   ├── admin/         # Admin-only APIs
│   │   │   │   ├── analytics/ # GET /api/admin/analytics
│   │   │   │   ├── disputes/  # GET/PUT /api/admin/disputes
│   │   │   │   ├── jobs/      # GET/PATCH/DELETE /api/admin/jobs
│   │   │   │   ├── reports/   # GET/PUT /api/admin/reports
│   │   │   │   └── users/     # GET/PATCH /api/admin/users
│   │   │   ├── contracts/     # Contract management
│   │   │   │   ├── [id]/
│   │   │   │   │   ├── milestones/  # Milestone CRUD
│   │   │   │   │   │   └── [mid]/   # Individual milestone actions
│   │   │   │   │   └── route.ts     # Contract details & status
│   │   │   │   └── route.ts         # List user contracts
│   │   │   ├── conversations/ # Messaging
│   │   │   │   ├── [id]/
│   │   │   │   │   └── messages/    # GET/POST messages
│   │   │   │   └── route.ts         # Create conversation
│   │   │   ├── freelancers/   # Freelancer profiles
│   │   │   │   ├── [id]/      # Public profile details
│   │   │   │   └── route.ts   # Browse freelancers
│   │   │   ├── jobs/          # Job management
│   │   │   │   ├── [id]/
│   │   │   │   │   ├── proposals/   # Job proposals
│   │   │   │   │   └── route.ts     # Job details
│   │   │   │   └── route.ts         # List/create jobs
│   │   │   ├── notifications/ # User notifications
│   │   │   ├── proposals/     # Proposal management
│   │   │   │   └── [id]/      # Accept/reject/withdraw
│   │   │   ├── reports/       # Fraud reports
│   │   │   ├── reviews/       # Reviews
│   │   │   ├── saved-jobs/    # Bookmarked jobs
│   │   │   └── transactions/  # Payment transactions
│   │   ├── globals.css        # Global styles (emerald theme)
│   │   ├── layout.tsx         # Root layout with metadata
│   │   └── page.tsx           # SPA entry point (client component)
│   ├── components/
│   │   ├── auth/              # Authentication components
│   │   │   ├── auth-provider.tsx   # Auth context & route protection
│   │   │   ├── login-form.tsx      # Login page
│   │   │   └── register-form.tsx   # Multi-step registration
│   │   ├── landing/           # Public landing page
│   │   │   └── landing-page.tsx
│   │   ├── layout/            # App shell & navigation
│   │   │   ├── app-shell.tsx       # Main SPA shell with view router
│   │   │   ├── sidebar.tsx         # Role-based sidebar navigation
│   │   │   └── top-nav.tsx         # Top navigation bar
│   │   ├── pages/             # All page components
│   │   │   ├── browse-freelancers.tsx
│   │   │   ├── client-dashboard.tsx
│   │   │   ├── contract-detail.tsx
│   │   │   ├── freelancer-dashboard.tsx
│   │   │   ├── freelancer-profile.tsx
│   │   │   ├── job-detail.tsx
│   │   │   ├── jobs-marketplace.tsx
│   │   │   ├── my-contracts.tsx
│   │   │   ├── my-jobs.tsx
│   │   │   ├── my-proposals.tsx
│   │   │   ├── post-job-form.tsx
│   │   │   ├── profile-editor.tsx
│   │   │   ├── reviews-page.tsx
│   │   │   ├── saved-jobs.tsx
│   │   │   └── settings-page.tsx
│   │   ├── shared/            # Shared/reusable components
│   │   └── ui/                # shadcn/ui components
│   ├── hooks/                 # Custom React hooks
│   ├── lib/                   # Utilities & configurations
│   │   ├── api-auth.ts        # API authentication middleware
│   │   ├── auth.ts            # NextAuth configuration
│   │   ├── auth-utils.ts      # Auth helper functions
│   │   └── db.ts              # Prisma client singleton
│   └── store/                 # Zustand state stores
│       ├── auth-store.ts      # Authentication state
│       └── navigation-store.ts # SPA navigation state
├── .env                       # Environment variables (not committed)
├── .env.example               # Example environment variables
├── package.json               # Dependencies and scripts
├── tailwind.config.ts         # Tailwind CSS configuration
├── tsconfig.json              # TypeScript configuration
└── README.md                  # This file
```

---

## Environment Variables

Create a `.env` file in the project root:

```env
# Database
# SQLite (development):
DATABASE_URL="file:./dev.db"

# PostgreSQL (production):
# DATABASE_URL="postgresql://studenthire_user:your_secure_password@localhost:5432/studenthire"

# Authentication
NEXTAUTH_URL="http://localhost:3000"
NEXTAUTH_SECRET="your-secret-key-here"
# Generate a secret with: openssl rand -base64 32
```

### Variable Reference

| Variable | Required | Default | Description |
|---|---|---|---|
| `DATABASE_URL` | Yes | — | Database connection string (SQLite or PostgreSQL) |
| `NEXTAUTH_URL` | Yes | `http://localhost:3000` | Base URL for NextAuth callbacks |
| `NEXTAUTH_SECRET` | Yes | — | Secret for signing JWT tokens (min 32 chars) |

---

## Demo Accounts

The seed script creates the following demo accounts. All passwords are `password123`.

| Role | Email | Password | Description |
|---|---|---|---|
| 🔑 **Admin** | `admin@studenthire.com` | `password123` | Platform administrator with full access |
| 🏢 **Client** | `client@demo.com` | `password123` | "TechCorp Solutions" — posted 5 demo jobs |
| 🎨 **Freelancer** | `freelancer@demo.com` | `password123` | "CreativeWolf42" — MIT CS student, web developer |
| 🏢 **Past Client** | `pastclient@demo.com` | `password123` | "DesignStudio Pro" — left a 4-star review |

### Seed Data Summary

| Data Type | Count | Details |
|---|---|---|
| Users | 4 | 1 Admin, 2 Clients, 1 Freelancer |
| Jobs | 6 | 4 Open, 1 In Progress, 1 Completed |
| Proposals | 4 | 2 Pending, 2 Accepted |
| Contracts | 2 | 1 Active (3 milestones), 1 Completed |
| Milestones | 3 | 1 Approved, 1 In Progress, 1 Pending |
| Messages | 5 | Between client and freelancer |
| Reviews | 2 | 5 stars + 4 stars |
| Notifications | 3 | For the freelancer |
| Transactions | 4 | 2 Escrow deposits, 2 Payment releases |
| Skills | 3 | React, Graphic Design, Video Editing |

---

## Security Features (Fraud Prevention)

StudentHire implements multiple layers of protection to ensure a safe marketplace environment:

### Identity Protection
- **Display Name System** — Users never see each other's real names or email addresses
- **Auto-Generated Names** — Random display names are assigned at registration (e.g., "CreativeWolf42")
- **Profile Separation** — User credentials are stored separately from public profiles

### Financial Security
- **Escrow Payments** — Clients deposit funds before work begins; freelancers are paid only upon milestone approval
- **Milestone-Based Releases** — Payments are released incrementally as work is approved
- **Platform Fee** — 10% fee on payment releases ensures platform sustainability
- **Transaction Logging** — All financial transactions are recorded with full audit trail

### User Safety
- **Report System** — Users can flag suspicious behavior (scam, harassment, spam, inappropriate content, payment disputes)
- **Admin Moderation** — All reports are reviewed by administrators with resolution tracking
- **Dispute Resolution** — Admin-mediated conflict resolution with refund/release capabilities
- **Account Suspension** — Admins can suspend accounts with documented reasons
- **User Verification** — Email verification and optional identity verification for enhanced trust

### Student Protection
- **Student Verification** — Special verification for student users (institution, major, graduation date)
- **Student Badge** — Verified students get a visible badge on their profiles
- **Student Filters** — Clients can specifically search for student freelancers

---

## API Documentation

All API routes require authentication (except public job listings and freelancer profiles). Use NextAuth session cookies or JWT tokens.

### Authentication

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `POST` | `/api/auth/register` | Public | Register new user (CLIENT or FREELANCER) |
| `GET` | `/api/auth/[...nextauth]` | Public | NextAuth handlers (signIn, signOut, session) |
| `POST` | `/api/auth/[...nextauth]` | Public | NextAuth credentials sign-in |

### Jobs

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/jobs` | Public | List jobs with filters (category, budget, status, search, pagination) |
| `POST` | `/api/jobs` | CLIENT | Create a new job listing |
| `GET` | `/api/jobs/[id]` | Public | Get job details with client profile |
| `PUT` | `/api/jobs/[id]` | CLIENT (owner) | Update job listing |
| `DELETE` | `/api/jobs/[id]` | CLIENT (owner) | Soft-delete job |

### Proposals

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/jobs/[id]/proposals` | CLIENT (owner) | List proposals for a job |
| `POST` | `/api/jobs/[id]/proposals` | FREELANCER | Submit a proposal |
| `PUT` | `/api/proposals/[id]` | CLIENT (owner) | Update proposal status (shortlist/accept/reject) |
| `PATCH` | `/api/proposals/[id]` | FREELANCER (owner) | Withdraw proposal |

### Contracts

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/contracts` | Authenticated | List user's contracts |
| `GET` | `/api/contracts/[id]` | Authenticated | Get contract details with milestones |
| `PUT` | `/api/contracts/[id]` | CLIENT or FREELANCER | Update contract status |

### Milestones

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/contracts/[id]/milestones` | Authenticated | List contract milestones |
| `POST` | `/api/contracts/[id]/milestones` | CLIENT (owner) | Create a milestone |
| `PUT` | `/api/contracts/[id]/milestones/[mid]` | Authenticated | Update milestone (submit work) |
| `PATCH` | `/api/contracts/[id]/milestones/[mid]` | CLIENT (owner) | Approve or reject milestone |

### Conversations & Messages

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/conversations` | Authenticated | List conversations |
| `POST` | `/api/conversations` | Authenticated | Create or find conversation |
| `GET` | `/api/conversations/[id]/messages` | Authenticated | Get conversation messages |
| `POST` | `/api/conversations/[id]/messages` | Authenticated | Send a message |

### Reviews

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `POST` | `/api/reviews` | Authenticated | Create a review (post-contract) |
| `GET` | `/api/reviews` | Public | Get reviews for a user |

### Freelancers

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/freelancers` | Public | Browse freelancers (search, filter, sort) |
| `GET` | `/api/freelancers/[id]` | Public | Get freelancer public profile |

### Transactions

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/transactions` | Authenticated | List transactions with financial summary |
| `POST` | `/api/transactions` | CLIENT | Create escrow deposit |

### Notifications

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/notifications` | Authenticated | List notifications (with unread count) |
| `PATCH` | `/api/notifications` | Authenticated | Mark notifications as read |

### Reports

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `POST` | `/api/reports` | Authenticated | File a report |
| `GET` | `/api/reports` | Authenticated | List filed reports |

### Saved Jobs

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/saved-jobs` | Authenticated | List saved/bookmarked jobs |
| `POST` | `/api/saved-jobs` | Authenticated | Save a job |
| `DELETE` | `/api/saved-jobs` | Authenticated | Remove a saved job |

### Admin (ADMIN only)

| Method | Endpoint | Auth | Description |
|---|---|---|---|
| `GET` | `/api/admin/users` | ADMIN | List all users (with search, filters) |
| `PATCH` | `/api/admin/users` | ADMIN | Suspend/verify users |
| `GET` | `/api/admin/jobs` | ADMIN | List all jobs |
| `PATCH` | `/api/admin/jobs` | ADMIN | Feature/remove jobs |
| `DELETE` | `/api/admin/jobs` | ADMIN | Delete jobs |
| `GET` | `/api/admin/reports` | ADMIN | List all reports |
| `PUT` | `/api/admin/reports` | ADMIN | Update report status with notes |
| `GET` | `/api/admin/disputes` | ADMIN | List disputed contracts |
| `PUT` | `/api/admin/disputes` | ADMIN | Resolve disputes (release/refund) |
| `GET` | `/api/admin/analytics` | ADMIN | Platform statistics dashboard |

---

## Database Schema

The platform uses 15 database models with 18 enums:

**Models:** User, FreelancerProfile, ClientProfile, Job, Proposal, Contract, Milestone, Conversation, Message, Review, Report, Transaction, Notification, SavedJob, Skill

**Enums:** UserRole, JobCategory, BudgetType, ProjectDuration, ExperienceLevel, EnglishLevel, ProposalStatus, ContractStatus, MilestoneStatus, MessageType, ReportType, ReportStatus, TransactionType, TransactionStatus, NotificationType, AvailabilityStatus, RelatedEntityType

See `prisma/schema.prisma` for the complete schema definition.

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -m 'Add my feature'`
4. Push to the branch: `git push origin feature/my-feature`
5. Submit a pull request

## License

This project is licensed under the MIT License.

---

<p align="center">
  Built with ❤️ for student freelancers everywhere
</p>
