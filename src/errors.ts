export type PptxErrorCode =
  /** Input is not a readable .pptx archive. */
  | "INVALID_FILE"
  /** Archive exceeds a configured safety limit (zip bomb protection). */
  | "LIMIT_EXCEEDED"
  /** Loading the file over the network failed. */
  | "FETCH_FAILED"
  /** A part of the document could not be rendered; rendering continued. */
  | "RENDER_WARNING"
  /** Export to .pptx/.odp failed. */
  | "EXPORT_FAILED"
  /** Operation aborted through an AbortSignal. */
  | "ABORTED";

/** Error type thrown (or reported through `onWarning`) by PPTXjs. */
export class PptxError extends Error {
  readonly code: PptxErrorCode;
  readonly cause?: unknown;
  /** Slide number (1-based) the error relates to, when known. */
  slide?: number;

  constructor(code: PptxErrorCode, message: string, cause?: unknown) {
    super(message);
    this.name = "PptxError";
    this.code = code;
    this.cause = cause;
  }
}
