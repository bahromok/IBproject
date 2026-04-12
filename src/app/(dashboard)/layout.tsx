'use client'

import { Web3Provider } from "@/components/providers/web3-provider"

export default function DashboardLayout({
    children,
}: {
    children: React.ReactNode
}) {
    return (
        <Web3Provider>
            {children}
        </Web3Provider>
    )
}
