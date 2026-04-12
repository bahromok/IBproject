'use server'

import { db } from '@/lib/db'
import { auth } from '@/lib/auth/server'
import { headers } from 'next/headers'

export async function switchUserRole(newRole: 'CLIENT' | 'FREELANCER') {
    const session = await auth.api.getSession({
        headers: await headers()
    })

    if (!session?.user) {
        throw new Error('Not authenticated')
    }

    const userId = session.user.id

    await db.user.update({
        where: { id: userId },
        data: { role: newRole }
    })

    return { success: true }
}
