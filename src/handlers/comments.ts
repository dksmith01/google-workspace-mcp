import type { drive_v3 } from "googleapis";
import { log, structuredResponse, validateArgs } from "../utils/index.js";
import type { ToolResponse } from "../utils/index.js";
import {
  ListCommentsSchema,
  ReplyToCommentSchema,
  ResolveCommentSchema,
} from "../schemas/index.js";

const COMMENT_FIELDS =
  "id, content, quotedFileContent(value), author(displayName), createdTime, modifiedTime, " +
  "resolved, replies(id, content, author(displayName), createdTime, action)";

interface CommentReply {
  id?: string;
  author?: string;
  content?: string;
  action?: string;
  createdTime?: string;
}

interface CommentSummary {
  id?: string;
  author?: string;
  content?: string;
  quotedText?: string;
  resolved: boolean;
  createdTime?: string;
  modifiedTime?: string;
  replies: CommentReply[];
}

function toCommentSummary(comment: drive_v3.Schema$Comment): CommentSummary {
  return {
    id: comment.id ?? undefined,
    author: comment.author?.displayName ?? undefined,
    content: comment.content ?? undefined,
    quotedText: comment.quotedFileContent?.value ?? undefined,
    resolved: comment.resolved ?? false,
    createdTime: comment.createdTime ?? undefined,
    modifiedTime: comment.modifiedTime ?? undefined,
    replies: (comment.replies ?? []).map((reply) => ({
      id: reply.id ?? undefined,
      author: reply.author?.displayName ?? undefined,
      content: reply.content ?? undefined,
      action: reply.action ?? undefined,
      createdTime: reply.createdTime ?? undefined,
    })),
  };
}

function formatComment(comment: CommentSummary, index: number): string {
  const status = comment.resolved ? "resolved" : "open";
  const lines = [
    `${index + 1}. [${status}] ${comment.author ?? "Unknown"} (${comment.createdTime ?? "unknown time"})`,
    `   ID: ${comment.id}`,
  ];
  if (comment.quotedText) lines.push(`   On: "${comment.quotedText}"`);
  lines.push(`   ${comment.content ?? ""}`);
  for (const reply of comment.replies) {
    const action = reply.action ? ` (${reply.action})` : "";
    lines.push(`   ↳ ${reply.author ?? "Unknown"}${action}: ${reply.content ?? ""}`);
  }
  return lines.join("\n");
}

export async function handleListComments(
  drive: drive_v3.Drive,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(ListCommentsSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  const allComments: drive_v3.Schema$Comment[] = [];
  let pageToken: string | undefined;
  do {
    const response = await drive.comments.list({
      fileId: data.fileId,
      fields: `nextPageToken, comments(${COMMENT_FIELDS})`,
      pageSize: 100,
      pageToken,
    });
    allComments.push(...(response.data.comments ?? []));
    pageToken = response.data.nextPageToken ?? undefined;
  } while (pageToken);

  const comments = allComments
    .filter((comment) => data.includeResolved || !comment.resolved)
    .map(toCommentSummary);

  log("Listed comments", { fileId: data.fileId, count: comments.length });

  if (comments.length === 0) {
    const scope = data.includeResolved ? "comments" : "open comments";
    return structuredResponse(`No ${scope} on this file.`, {
      fileId: data.fileId,
      comments: [],
      total: 0,
    });
  }

  const text = comments.map(formatComment).join("\n\n");
  return structuredResponse(`${comments.length} comment(s):\n\n${text}`, {
    fileId: data.fileId,
    comments,
    total: comments.length,
  });
}

export async function handleReplyToComment(
  drive: drive_v3.Drive,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(ReplyToCommentSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  const response = await drive.replies.create({
    fileId: data.fileId,
    commentId: data.commentId,
    fields: "id, content, createdTime",
    requestBody: { content: data.content },
  });

  log("Replied to comment", { fileId: data.fileId, commentId: data.commentId });

  return structuredResponse(`Replied to comment ${data.commentId}`, {
    commentId: data.commentId,
    replyId: response.data.id!,
  });
}

export async function handleResolveComment(
  drive: drive_v3.Drive,
  args: unknown,
): Promise<ToolResponse> {
  const validation = validateArgs(ResolveCommentSchema, args);
  if (!validation.success) return validation.response;
  const data = validation.data;

  const response = await drive.replies.create({
    fileId: data.fileId,
    commentId: data.commentId,
    fields: "id, action",
    requestBody: {
      action: "resolve",
      ...(data.content && { content: data.content }),
    },
  });

  log("Resolved comment", { fileId: data.fileId, commentId: data.commentId });

  return structuredResponse(`Resolved comment ${data.commentId}`, {
    commentId: data.commentId,
    replyId: response.data.id!,
    resolved: true,
  });
}
