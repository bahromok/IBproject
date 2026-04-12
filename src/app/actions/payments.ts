'use server'

import { db } from '@/lib/db'
import { auth } from '@/lib/auth/server'
import { headers } from 'next/headers'

import { stripe } from '@/lib/stripe'

export async function recordCryptoTransaction(milestoneId: string, txHash: string, fromAddress: string) {
    const session = await auth.api.getSession({
        headers: await headers()
    })

    if (!session?.user) throw new Error('Not authenticated')

    const milestone = await db.milestone.findUnique({
        where: { id: milestoneId },
        include: { contract: true }
    })

    if (!milestone) throw new Error('Milestone not found')
    if (milestone.contract.clientId !== session.user.id) throw new Error('Unauthorized')

    const platformFee = milestone.amount * 0.05
    const netAmount = milestone.amount * 0.95

    await db.$transaction([
        db.milestone.update({
            where: { id: milestoneId },
            data: { status: 'IN_PROGRESS' } // Funded
        }),
        db.transaction.create({
            data: {
                milestoneId,
                contractId: milestone.contractId,
                fromUserId: milestone.contract.clientId,
                toUserId: milestone.contract.freelancerId,
                amount: milestone.amount,
                platformFee,
                netAmount,
                type: 'ESCROW_DEPOSIT',
                status: 'COMPLETED',
                description: `Crypto Transaction Hash: ${txHash}. 5% Commission for Hamkorbank (5614681814216473)`
            }
        })
    ])

    return { success: true }
}

export async function getStripeConnectUrl() {
    const session = await auth.api.getSession({
        headers: await headers()
    })

    if (!session?.user) throw new Error('Not authenticated')

    // 1. Get the user from DB to check for existing Stripe ID
    const user = await db.user.findUnique({
        where: { id: session.user.id }
    })

    if (!user) throw new Error('User not found')

    // @ts-ignore
    let accountId = user.stripeConnectId

    // 2. Create a Connect account if one doesn't exist
    if (!accountId) {
        const account = await stripe.accounts.create({
            type: 'express',
            email: user.email,
            capabilities: {
                card_payments: { requested: true },
                transfers: { requested: true },
            },
        });

        accountId = account.id

        // Save the ID to the user
        // @ts-ignore
        await db.user.update({
            where: { id: user.id },
            // @ts-ignore
            data: { stripeConnectId: accountId }
        })
    }

    // 3. Create the Account Link
    const host = (await headers()).get('host');
    const protocol = (await headers()).get('x-forwarded-proto') || 'http';
    const baseUrl = `${protocol}://${host}`;

    const accountLink = await stripe.accountLinks.create({
        account: accountId,
        refresh_url: `${baseUrl}/dashboard/settings?tab=payments&stripe=refresh`,
        return_url: `${baseUrl}/dashboard/settings?tab=payments&stripe=success`,
        type: 'account_onboarding',
    });

    return { url: accountLink.url };
}
