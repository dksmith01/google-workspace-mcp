import { z } from "zod";

export const ListCommentsSchema = z.object({
  fileId: z.string().min(1, "File ID is required"),
  includeResolved: z.boolean().optional().default(false),
});

export const ReplyToCommentSchema = z.object({
  fileId: z.string().min(1, "File ID is required"),
  commentId: z.string().min(1, "Comment ID is required"),
  content: z.string().min(1, "Reply content is required"),
});

export const ResolveCommentSchema = z.object({
  fileId: z.string().min(1, "File ID is required"),
  commentId: z.string().min(1, "Comment ID is required"),
  content: z.string().optional(),
});

// Type exports
export type ListCommentsInput = z.infer<typeof ListCommentsSchema>;
export type ReplyToCommentInput = z.infer<typeof ReplyToCommentSchema>;
export type ResolveCommentInput = z.infer<typeof ResolveCommentSchema>;
