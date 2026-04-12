import { NextResponse } from 'next/server';
import { auth } from '@/lib/auth/server';
import { stripe } from '@/lib/stripe';
import { db } from '@/lib/db';

export async function POST(req: Request) {
    try {
        const session = await auth.api.getSession({ headers: req.headers });
        if (!session || !session.user) {
            return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
        }

        const { milestoneId } = await req.json();

        const milestone = await db.milestone.findUnique({
            where: { id: milestoneId },
            include: {
                contract: {
                    include: {
                        client: true,
                    }
                }
            }
        });

        if (!milestone) {
            return NextResponse.json({ error: 'Milestone not found' }, { status: 404 });
        }

        if (milestone.contract.client.id !== session.user.id) {
            return NextResponse.json({ error: 'Unauthorized: You are not the client for this contract' }, { status: 403 });
        }

        if (milestone.status !== 'PENDING') {
            return NextResponse.json({ error: 'Milestone is already funded or in progress' }, { status: 400 });
        }

        const host = req.headers.get('host');
        const protocol = req.headers.get('x-forwarded-proto') || 'http';
        const baseUrl = `${protocol}://${host}`;

        const stripeSession = await stripe.checkout.sessions.create({
            payment_method_types: ['card'],
            line_items: [
                {
                    price_data: {
                        currency: 'usd',
                        product_data: {
                            name: `Milestone: ${milestone.title}`,
                            description: `Funding for milestone in contract: ${milestone.contract.title}`,
                        },
                        unit_amount: Math.round(milestone.amount * 100),
                    },
                    quantity: 1,
                },
            ],
            mode: 'payment',
            success_url: `${baseUrl}/dashboard/contracts/${milestone.contract.id}?payment=success`,
            cancel_url: `${baseUrl}/dashboard/contracts/${milestone.contract.id}?payment=cancelled`,
            metadata: {
                milestoneId: milestone.id,
                contractId: milestone.contract.id,
                clientId: session.user.id,
                platformFee: Math.round(milestone.amount * 0.05 * 100) / 100, // 5%
            },
            customer_email: session.user.email,
        });

        return NextResponse.json({ url: stripeSession.url });
    } catch (error: any) {
        console.error('[STRIPE_CHECKOUT_ERROR]:', error);
        return NextResponse.json({ error: error.message || 'Internal Server Error' }, { status: 500 });
    }
}
