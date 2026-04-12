'use client'

import { useState, useEffect, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { useAuthStore } from '@/store/auth-store'
import { Input } from '@/components/ui/input'
import { Button } from '@/components/ui/button'
import { Avatar, AvatarFallback, AvatarImage } from '@/components/ui/avatar'
import { ScrollArea } from '@/components/ui/scroll-area'
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog'
import { Skeleton } from '@/components/ui/skeleton'
import { SafetyBanner } from '@/components/shared/safety-banner'
import { EmptyState } from '@/components/shared/empty-state'
import {
  MessageSquare,
  Search,
  Send,
  Paperclip,
  Plus,
  FileText,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { formatDistanceToNow } from 'date-fns'
import { authFetch } from '@/lib/api-fetch'

interface Conversation {
  id: string
  otherUser: { id: string; displayName: string; avatar: string | null }
  contract: { id: string; title: string } | null
  unreadCount: number
  lastMessageAt: string
  createdAt: string
}

interface LastMessageMap {
  [convId: string]: string
}

export function MessagesInbox() {
  const { navigate, pageParams } = useNavigationStore()
  const { user } = useAuthStore()
  const [conversations, setConversations] = useState<Conversation[]>([])
  const [lastMessageMap, setLastMessageMap] = useState<LastMessageMap>({})
  const [loading, setLoading] = useState(true)
  const [loadingMessages, setLoadingMessages] = useState(false)
  const [search, setSearch] = useState('')
  const [showNewMsgDialog, setShowNewMsgDialog] = useState(false)
  const [newMsgUserId, setNewMsgUserId] = useState('')
  const [newMsgSearch, setNewMsgSearch] = useState('')
  const [searchedUsers, setSearchedUsers] = useState<any[]>([])
  const [searchingUsers, setSearchingUsers] = useState(false)
  const [creatingConv, setCreatingConv] = useState(false)

  const fetchConversations = useCallback(async () => {
    setLoading(true)
    try {
      const res = await authFetch('/api/conversations')
      if (res.ok) {
        const data = await res.json()
        setConversations(data.conversations || [])
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    fetchConversations()
  }, [fetchConversations])

  // Fetch last messages for each conversation
  useEffect(() => {
    const fetchLastMessages = async () => {
      setLoadingMessages(true)
      const map: LastMessageMap = {}
      for (const conv of conversations) {
        try {
          const res = await authFetch(`/api/conversations/${conv.id}/messages?limit=1`)
          if (res.ok) {
            const data = await res.json()
            const msgs = data.messages || []
            if (msgs.length > 0) {
              map[conv.id] = msgs[msgs.length - 1].content
            }
          }
        } catch {
          // skip
        }
      }
      setLastMessageMap(map)
      setLoadingMessages(false)
    }
    if (conversations.length > 0) {
      fetchLastMessages()
    }
  }, [conversations.length])

  const handleCreateConversation = async () => {
    if (!newMsgUserId || creatingConv) return
    setCreatingConv(true)
    try {
      const res = await authFetch('/api/conversations', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ user2Id: newMsgUserId }),
      })
      if (res.ok) {
        const data = await res.json()
        const convId = data.conversation?.id
        if (convId) {
          setShowNewMsgDialog(false)
          setNewMsgUserId('')
          fetchConversations()
          navigate('messages/conversation', { id: convId })
        }
      }
    } catch {
      // Silently handle
    } finally {
      setCreatingConv(false)
    }
  }

  const handleSearchUsers = async (query: string) => {
    setNewMsgSearch(query)
    if (query.length < 2) {
      setSearchedUsers([])
      return
    }
    setSearchingUsers(true)
    try {
      const res = await authFetch(`/api/freelancers?search=${encodeURIComponent(query)}&limit=10`)
      if (res.ok) {
        const data = await res.json()
        setSearchedUsers(data.freelancers || data.data || [])
      }
    } catch {
      // Silently handle
    } finally {
      setSearchingUsers(false)
    }
  }

  const filteredConversations = conversations.filter(c =>
    c.otherUser.displayName.toLowerCase().includes(search.toLowerCase())
  )

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="space-y-4"
    >
      <SafetyBanner />

      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Messages</h1>
          <p className="text-slate-500 mt-1">
            {!loading && conversations.length > 0
              ? `${conversations.length} conversation${conversations.length !== 1 ? 's' : ''}`
              : 'Your conversations'}
          </p>
        </div>
        <Dialog open={showNewMsgDialog} onOpenChange={setShowNewMsgDialog}>
          <DialogTrigger asChild>
            <Button size="sm" className="gap-1.5 bg-emerald-500 hover:bg-emerald-600">
              <Plus className="w-4 h-4" />
              New Message
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Start a New Conversation</DialogTitle>
            </DialogHeader>
            <div className="space-y-4 mt-4">
              <Input
                placeholder="Search by display name..."
                value={newMsgSearch}
                onChange={e => handleSearchUsers(e.target.value)}
              />
              <ScrollArea className="max-h-64">
                {searchingUsers ? (
                  <div className="space-y-2 p-2">
                    {[1, 2, 3].map(i => (
                      <Skeleton key={i} className="h-12 w-full rounded-lg" />
                    ))}
                  </div>
                ) : searchedUsers.length > 0 ? (
                  <div className="space-y-1">
                    {searchedUsers.map((u: any) => (
                      <button
                        key={u.id || u.userId}
                        onClick={() => setNewMsgUserId(u.id || u.userId)}
                        className={cn(
                          'w-full flex items-center gap-3 p-2.5 rounded-lg text-left transition-colors',
                          newMsgUserId === (u.id || u.userId)
                            ? 'bg-emerald-50 border border-emerald-200'
                            : 'hover:bg-slate-50'
                        )}
                      >
                        <Avatar className="h-8 w-8">
                          {u.avatar && <AvatarImage src={u.avatar} />}
                          <AvatarFallback className="bg-emerald-100 text-emerald-700 text-xs">
                            {(u.displayName || 'U').slice(0, 2).toUpperCase()}
                          </AvatarFallback>
                        </Avatar>
                        <div>
                          <div className="text-sm font-medium text-slate-900">{u.displayName}</div>
                          <div className="text-xs text-slate-500">{u.title || 'Freelancer'}</div>
                        </div>
                      </button>
                    ))}
                  </div>
                ) : newMsgSearch.length >= 2 ? (
                  <p className="text-sm text-slate-500 text-center py-4">No users found</p>
                ) : (
                  <p className="text-sm text-slate-400 text-center py-4">Type at least 2 characters to search</p>
                )}
              </ScrollArea>
              <Button
                onClick={handleCreateConversation}
                disabled={!newMsgUserId || creatingConv}
                className="w-full bg-emerald-500 hover:bg-emerald-600"
              >
                {creatingConv ? 'Creating...' : 'Start Conversation'}
              </Button>
            </div>
          </DialogContent>
        </Dialog>
      </div>

      {/* Search */}
      <div className="relative">
        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
        <Input
          placeholder="Search conversations..."
          className="pl-9 h-10"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
      </div>

      {/* Conversation List */}
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        {loading ? (
          <div className="space-y-0 p-4">
            {[1, 2, 3, 4, 5].map(i => (
              <div key={i} className="flex items-center gap-3 py-3">
                <Skeleton className="h-10 w-10 rounded-full shrink-0" />
                <div className="flex-1 min-w-0 space-y-1.5">
                  <Skeleton className="h-4 w-32" />
                  <Skeleton className="h-3 w-48" />
                </div>
                <Skeleton className="h-3 w-12 shrink-0" />
              </div>
            ))}
          </div>
        ) : filteredConversations.length === 0 ? (
          <EmptyState
            icon={MessageSquare}
            title={search ? 'No conversations found' : 'No conversations yet'}
            description={search
              ? 'Try a different search term.'
              : 'Click "New Message" to start a conversation with someone.'}
          />
        ) : (
          <div className="divide-y divide-slate-100 max-h-[calc(100vh-22rem)] overflow-y-auto">
            {filteredConversations.map((conv) => (
              <button
                key={conv.id}
                onClick={() => navigate('messages/conversation', { id: conv.id })}
                className="w-full flex items-center gap-3 p-4 text-left hover:bg-slate-50 transition-colors"
              >
                <div className="relative shrink-0">
                  <Avatar className="h-11 w-11">
                    {conv.otherUser.avatar && (
                      <AvatarImage src={conv.otherUser.avatar} alt={conv.otherUser.displayName} />
                    )}
                    <AvatarFallback className="bg-emerald-100 text-emerald-700 text-sm font-semibold">
                      {conv.otherUser.displayName.slice(0, 2).toUpperCase()}
                    </AvatarFallback>
                  </Avatar>
                  {conv.unreadCount > 0 && (
                    <span className="absolute -top-0.5 -right-0.5 w-5 h-5 rounded-full bg-emerald-500 text-white text-[10px] font-bold flex items-center justify-center">
                      {conv.unreadCount > 9 ? '9+' : conv.unreadCount}
                    </span>
                  )}
                </div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-sm font-semibold text-slate-900 truncate">
                      {conv.otherUser.displayName}
                    </span>
                    <span className="text-xs text-slate-400 shrink-0">
                      {conv.lastMessageAt
                        ? formatDistanceToNow(new Date(conv.lastMessageAt), { addSuffix: false })
                        : ''}
                    </span>
                  </div>
                  <div className="flex items-center gap-1 mt-0.5">
                    <p className="text-xs text-slate-500 truncate flex-1">
                      {loadingMessages ? (
                        <Skeleton className="h-3 w-40 inline-block" />
                      ) : (
                        lastMessageMap[conv.id] || 'No messages yet'
                      )}
                    </p>
                  </div>
                  {conv.contract && (
                    <div className="flex items-center gap-1 mt-1">
                      <FileText className="w-3 h-3 text-slate-400" />
                      <span className="text-[11px] text-slate-400 truncate">{conv.contract.title}</span>
                    </div>
                  )}
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </motion.div>
  )
}
