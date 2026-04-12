'use client'

import { useState, useEffect, useRef, useCallback } from 'react'
import { motion } from 'framer-motion'
import { useNavigationStore } from '@/store/navigation-store'
import { Button } from '@/components/ui/button'
import { Avatar, AvatarFallback, AvatarImage } from '@/components/ui/avatar'
import { ScrollArea } from '@/components/ui/scroll-area'
import { Skeleton } from '@/components/ui/skeleton'
import { EmptyState } from '@/components/shared/empty-state'
import {
  ArrowLeft,
  Send,
  Paperclip,
  ExternalLink,
  FileText,
  MessageSquare,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { formatDistanceToNow } from 'date-fns'
import { authFetch } from '@/lib/api-fetch'

interface Message {
  id: string
  content: string
  messageType: string
  fileUrl: string | null
  fileName: string | null
  isOwnMessage: boolean
  sender: { id: string; displayName: string; avatar: string | null }
  createdAt: string
}

interface ConversationInfo {
  id: string
  otherUser: { id: string; displayName: string; avatar: string | null }
  contract: { id: string; title: string } | null
}

export function MessagesConversation() {
  const { navigate, pageParams, goBack } = useNavigationStore()
  const convId = pageParams?.id
  const [messages, setMessages] = useState<Message[]>([])
  const [conversation, setConversation] = useState<ConversationInfo | null>(null)
  const [newMessage, setNewMessage] = useState('')
  const [sending, setSending] = useState(false)
  const [loading, setLoading] = useState(true)
  const messagesEndRef = useRef<HTMLDivElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)

  const fetchData = useCallback(async () => {
    if (!convId) return
    setLoading(true)
    try {
      const [convRes, msgRes] = await Promise.all([
        authFetch('/api/conversations'),
        authFetch(`/api/conversations/${convId}/messages?limit=100`),
      ])
      if (convRes.ok) {
        const convData = await convRes.json()
        const conv = (convData.conversations || []).find((c: any) => c.id === convId)
        if (conv) setConversation(conv)
      }
      if (msgRes.ok) {
        const msgData = await msgRes.json()
        setMessages(msgData.messages || [])
      }
    } catch {
      // Silently handle
    } finally {
      setLoading(false)
    }
  }, [convId])

  useEffect(() => {
    fetchData()
    const interval = setInterval(fetchData, 5000)
    return () => clearInterval(interval)
  }, [fetchData])

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages])

  const handleSend = async () => {
    if (!newMessage.trim() || !convId || sending) return
    setSending(true)
    try {
      const res = await authFetch(`/api/conversations/${convId}/messages`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ content: newMessage.trim() }),
      })
      if (res.ok) {
        setNewMessage('')
        if (textareaRef.current) textareaRef.current.style.height = 'auto'
        fetchData()
      }
    } catch {
      // Silently handle
    } finally {
      setSending(false)
    }
  }

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSend()
    }
  }

  const handleTextareaInput = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setNewMessage(e.target.value)
    const textarea = e.target
    textarea.style.height = 'auto'
    textarea.style.height = Math.min(textarea.scrollHeight, 150) + 'px'
  }

  if (!convId) {
    return (
      <div className="flex items-center justify-center h-[calc(100vh-12rem)]">
        <EmptyState
          icon={MessageSquare}
          title="No conversation selected"
          description="Select a conversation from your inbox to start messaging."
          actionLabel="Back to Messages"
          onAction={() => navigate('messages')}
        />
      </div>
    )
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      animate={{ opacity: 1, y: 0 }}
      className="flex flex-col h-[calc(100vh-12rem)] bg-white rounded-xl border border-slate-200 overflow-hidden"
    >
      {/* Header */}
      <div className="flex items-center gap-3 p-4 border-b border-slate-200 bg-white shrink-0">
        <Button variant="ghost" size="icon" className="h-8 w-8 shrink-0" onClick={() => navigate('messages')}>
          <ArrowLeft className="w-4 h-4" />
        </Button>
        {conversation ? (
          <>
            <Avatar className="h-10 w-10">
              {conversation.otherUser.avatar && (
                <AvatarImage src={conversation.otherUser.avatar} />
              )}
              <AvatarFallback className="bg-emerald-100 text-emerald-700 font-semibold">
                {conversation.otherUser.displayName.slice(0, 2).toUpperCase()}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1 min-w-0">
              <h3 className="text-sm font-semibold text-slate-900 truncate">
                {conversation.otherUser.displayName}
              </h3>
              {conversation.contract && (
                <div className="flex items-center gap-1">
                  <FileText className="w-3 h-3 text-slate-400" />
                  <span className="text-xs text-slate-500 truncate">{conversation.contract.title}</span>
                </div>
              )}
            </div>
            <Button
              variant="ghost"
              size="sm"
              className="h-8 text-xs text-slate-500"
              onClick={() => navigate('freelancers/detail', { id: conversation.otherUser.id })}
            >
              <ExternalLink className="w-3 h-3 mr-1" />
              Profile
            </Button>
          </>
        ) : (
          <div className="flex items-center gap-3">
            <Skeleton className="h-10 w-10 rounded-full" />
            <Skeleton className="h-5 w-36" />
          </div>
        )}
      </div>

      {/* Messages */}
      <ScrollArea className="flex-1 p-4">
        {loading ? (
          <div className="space-y-3">
            {[1, 2, 3, 4, 5].map(i => (
              <Skeleton key={i} className={cn('h-12 rounded-xl max-w-xs', i % 2 === 0 ? 'ml-auto' : '')} />
            ))}
          </div>
        ) : messages.length === 0 ? (
          <div className="flex items-center justify-center h-full min-h-[300px]">
            <EmptyState
              icon={MessageSquare}
              title="No messages yet"
              description="Send the first message to start the conversation."
            />
          </div>
        ) : (
          <div className="space-y-3">
            {messages.map((msg) => (
              <motion.div
                key={msg.id}
                initial={{ opacity: 0, y: 5 }}
                animate={{ opacity: 1, y: 0 }}
                className={cn('flex gap-2', msg.isOwnMessage ? 'justify-end' : 'justify-start')}
              >
                {!msg.isOwnMessage && (
                  <Avatar className="h-7 w-7 shrink-0 mt-auto">
                    {msg.sender.avatar && <AvatarImage src={msg.sender.avatar} />}
                    <AvatarFallback className="bg-emerald-100 text-emerald-700 text-[10px]">
                      {msg.sender.displayName.slice(0, 2).toUpperCase()}
                    </AvatarFallback>
                  </Avatar>
                )}
                <div className={cn('max-w-[80%] md:max-w-[70%]', msg.messageType === 'SYSTEM' && 'max-w-full mx-auto')}>
                  {msg.messageType === 'SYSTEM' ? (
                    <div className="text-center py-2">
                      <span className="text-xs text-slate-400 bg-slate-50 px-3 py-1 rounded-full">
                        {msg.content}
                      </span>
                    </div>
                  ) : (
                    <div>
                      <div
                        className={cn(
                          'rounded-2xl px-3.5 py-2.5 text-sm',
                          msg.isOwnMessage
                            ? 'bg-emerald-500 text-white rounded-br-md'
                            : 'bg-slate-100 text-slate-900 rounded-bl-md'
                        )}
                      >
                        {msg.fileUrl && (
                          <div className={cn(
                            'flex items-center gap-2 mb-2 p-2 rounded-lg',
                            msg.isOwnMessage ? 'bg-white/10' : 'bg-white'
                          )}>
                            <Paperclip className="w-3.5 h-3.5 shrink-0" />
                            <span className="text-xs truncate">{msg.fileName || 'Attachment'}</span>
                            <a
                              href={msg.fileUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="text-xs underline shrink-0"
                            >
                              Download
                            </a>
                          </div>
                        )}
                        <p className="whitespace-pre-wrap break-words">{msg.content}</p>
                      </div>
                      <div className={cn(
                        'text-[10px] text-slate-400 mt-1 px-1',
                        msg.isOwnMessage ? 'text-right' : 'text-left'
                      )}>
                        {formatDistanceToNow(new Date(msg.createdAt), { addSuffix: true })}
                      </div>
                    </div>
                  )}
                </div>
              </motion.div>
            ))}
            <div ref={messagesEndRef} />
          </div>
        )}
      </ScrollArea>

      {/* Input */}
      <div className="p-4 border-t border-slate-200 bg-white shrink-0">
        <div className="flex items-end gap-2">
          <Button variant="ghost" size="icon" className="h-9 w-9 shrink-0 text-slate-400">
            <Paperclip className="w-4 h-4" />
          </Button>
          <div className="flex-1 relative">
            <textarea
              ref={textareaRef}
              value={newMessage}
              onChange={handleTextareaInput}
              onKeyDown={handleKeyDown}
              placeholder="Type a message..."
              rows={1}
              className="w-full resize-none rounded-xl border border-slate-200 bg-slate-50 px-4 py-2.5 text-sm text-slate-900 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all max-h-[150px]"
            />
          </div>
          <Button
            size="icon"
            className="h-9 w-9 shrink-0 bg-emerald-500 hover:bg-emerald-600"
            onClick={handleSend}
            disabled={!newMessage.trim() || sending}
          >
            <Send className="w-4 h-4" />
          </Button>
        </div>
        <p className="text-[10px] text-slate-400 mt-1.5 text-center">
          Press Enter to send, Shift+Enter for new line
        </p>
      </div>
    </motion.div>
  )
}
