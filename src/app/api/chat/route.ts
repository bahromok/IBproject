import { NextRequest, NextResponse } from 'next/server';
import { runAutonomousAgent, HistoryMessage } from '@/lib/autonomous-agent';

export async function POST(request: NextRequest) {
  try {
    const { message, history } = await request.json();

    if (!message) {
      return NextResponse.json({ error: 'Message is required' }, { status: 400 });
    }

    const userMessageId = 'user-' + Date.now();
    const assistantMessageId = 'assistant-' + Date.now();
    const createdAt = new Date().toISOString();

    const conversationHistory: HistoryMessage[] = Array.isArray(history)
      ? history.map((h: any) => ({
          role: h.role || 'user',
          content: h.content || '',
          filenames: h.filenames || [],
        }))
      : [];

    // Abort support
    const abortController = new AbortController();
    let clientDisconnected = false;

    // Listen for client disconnect
    request.signal.addEventListener('abort', () => {
      clientDisconnected = true;
      abortController.abort();
    });

    const encoder = new TextEncoder();
    const stream = new ReadableStream({
      async start(controller) {
        try {
          const send = (data: any) => {
            if (!clientDisconnected) {
              controller.enqueue(encoder.encode('data: ' + JSON.stringify(data) + '\n\n'));
            }
          };

          send({ type: 'status', status: 'Starting...', progress: 0 });

          const result = await runAutonomousAgent(
            message,
            (status, progress, thinking) => {
              send({ type: 'progress', status, progress, thinking });
            },
            conversationHistory,
            abortController.signal
          );

          if (clientDisconnected) {
            controller.close();
            return;
          }

          const filenames = [...new Set(
            result.results
              .filter(r => r.filename)
              .map(r => r.filename)
          )];

          send({
            type: 'complete',
            thinking: result.thinking,
            message: result.finalMessage,
            results: result.results,
            isChat: result.isChat,
            userMessage: {
              id: userMessageId,
              role: 'user',
              content: message,
              createdAt: createdAt,
            },
            assistantMessage: {
              id: assistantMessageId,
              role: 'assistant',
              content: result.finalMessage,
              actions: result.results.map(r => r.message),
              filenames: filenames,
              createdAt: new Date().toISOString(),
            }
          });

          controller.close();
        } catch (error) {
          if (clientDisconnected) {
            controller.close();
            return;
          }
          controller.enqueue(encoder.encode('data: ' + JSON.stringify({
            type: 'error',
            error: error instanceof Error ? error.message : 'Unknown error'
          }) + '\n\n'));
          controller.close();
        }
      }
    });

    return new Response(stream, {
      headers: {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
      },
    });
  } catch (error) {
    console.error('Chat API error:', error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : 'Internal server error' },
      { status: 500 }
    );
  }
}

export async function GET() {
  return NextResponse.json({ messages: [] });
}

export async function DELETE() {
  return NextResponse.json({ success: true, message: 'Chat history cleared' });
}
