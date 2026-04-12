// ============================================================================
// StudentHire - Database Seed Script (Admin Only)
// ============================================================================
// Creates a single admin account for the StudentHire freelance marketplace.
// Run with: npx prisma db seed  or  bun run prisma/seed.ts
// ============================================================================

import { PrismaClient } from '@prisma/client';
import { hash } from 'bcryptjs';

const prisma = new PrismaClient();

async function main() {
  console.log('🌱 Seeding StudentHire database...\n');

  // ---------------------------------------------------------------------------
  // Clean existing data (reverse dependency order)
  // ---------------------------------------------------------------------------
  console.log('🧹 Cleaning existing data...');
  await prisma.fraudAlert.deleteMany();
  await prisma.notification.deleteMany();
  await prisma.transaction.deleteMany();
  await prisma.message.deleteMany();
  await prisma.savedJob.deleteMany();
  await prisma.review.deleteMany();
  await prisma.report.deleteMany();
  await prisma.milestone.deleteMany();
  await prisma.conversation.deleteMany();
  await prisma.contract.deleteMany();
  await prisma.proposal.deleteMany();
  await prisma.job.deleteMany();
  await prisma.freelancerProfile.deleteMany();
  await prisma.clientProfile.deleteMany();
  await prisma.skill.deleteMany();
  await prisma.user.deleteMany();
  console.log('✅ Existing data cleaned.\n');

  // ---------------------------------------------------------------------------
  // Create Admin Account
  // ---------------------------------------------------------------------------
  console.log('👤 Creating admin account...');

  const adminPasswordHash = await hash('password123!', 10);

  const admin = await prisma.user.create({
    data: {
      email: 'bahrommurzohamidow@gmail.com',
      passwordHash: adminPasswordHash,
      role: 'ADMIN',
      isVerified: true,
      isEmailVerified: true,
      lastLoginAt: new Date(),
    },
  });

  console.log(`   ✅ Admin: ${admin.email}`);
  console.log('');
  console.log('═══════════════════════════════════════════════════════════════');
  console.log('  🌱 StudentHire Database Seeded Successfully!');
  console.log('═══════════════════════════════════════════════════════════════');
  console.log('');
  console.log('  📊 Seed Summary:');
  console.log('  ───────────────────────────────────────────────────────────');
  console.log('  👤 Admin users: 1');
  console.log('');
  console.log('  🔐 Admin Credentials:');
  console.log('  ───────────────────────────────────────────────────────────');
  console.log('  Email:    bahrommurzohamidow@gmail.com');
  console.log('  Password: password123!');
  console.log('');
  console.log('═══════════════════════════════════════════════════════════════\n');
}

main()
  .catch((e) => {
    console.error('❌ Seed failed:', e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
