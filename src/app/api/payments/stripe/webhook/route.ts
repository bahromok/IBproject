import { NextResponse } from 'next/server';
import { stripe } from '@/lib/stripe';
import { db } from '@/lib/db';
import { headers } from 'next/headers';

const endpointSecret = process.env.STRIPE_WEBHOOK_SECRET;

export async function POST(req: Request) {
    const body = await req.text();
    const headersList = headers();
    const sig = (await headersList).get('stripe-signature') as string;

    let event;

    try {
        if (!sig || !endpointSecret) {
            console.error('Missing Stripe signature or webhook secret');
            return NextResponse.json({ error: 'Webhook Secret Setup Required' }, { status: 400 });
        }
        event = stripe.webhooks.constructEvent(body, sig, endpointSecret);
    } catch (err: any) {
        console.error(`Webhook Error: ${err.message}`);
        return NextResponse.json({ error: `Webhook Error: ${err.message}` }, { status: 400 });
    }

    const session = event.data.object as any;

    if (event.type === 'checkout.session.completed') {
        const milestoneId = session.metadata?.milestoneId;

        if (milestoneId) {
            try {
                await db.$transaction(async (tx) => {
                    // Update milestone status to IN_PROGRESS (funded)
                    const mil = await tx.milestone.update({
                        where: { id: milestoneId },
                        data: { status: 'IN_PROGRESS' },
                        include: { contract: true }
                    });

                    const amount = session.amount_total / 100;
                    const platformFee = amount * 0.05;
                    const netAmount = amount * 0.95;

                    // Log the transaction
                    await tx.transaction.create({
                        data: {
                            contractId: session.metadata.contractId,
                            milestoneId: milestoneId,
                            amount: amount,
                            platformFee,
                            netAmount,
                            fromUserId: session.metadata.clientId,
                            toUserId: mil.contract.freelancerId,
                            type: 'ESCROW_DEPOSIT',
                            status: 'COMPLETED',
                            description: `Stripe Payment ID: ${session.payment_intent}. 5% Commission for Hamkorbank (5614681814216473)`,
                        }
                    });
                });

                console.log(`Successfully processed payment for milestone ${milestoneId}`);
            } catch (error) {
                console.error(`Error updating milestone ${milestoneId} on payment completion:`, error);
                return NextResponse.json({ error: 'DB Update Failed' }, { status: 500 });
            }
        }
    }

    return NextResponse.json({ received: true });
}
