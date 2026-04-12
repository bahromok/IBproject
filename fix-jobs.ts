import { PrismaClient } from '@prisma/client'
const prisma = new PrismaClient()

async function main() {
    const result = await prisma.job.updateMany({
        where: { status: 'open' },
        data: { status: 'OPEN' }
    })
    console.log(`Updated ${result.count} jobs to OPEN status.`)
}

main().catch(console.error).finally(() => prisma.$disconnect())
