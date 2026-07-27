/// <reference types="vite/client" />

export type PreviewPayload = {
  title: string;
  isFullMode: boolean;
  schema?: string;
  table: (string | number | null)[][];
  layout: {
    width: number;
    heightMin: number;
    gapV: number;
    gapH: number;
    baseLeft: number;
    baseTop: number;
  };
  meta?: {
    nodeCount?: number;
    colCount?: number;
    live?: boolean;
    embedded?: boolean;
    watch?: Record<string, unknown>;
  };
};

export type PreviewHostApi = {
  confirm: () => Promise<void> | void;
  cancel: () => Promise<void> | void;
};

declare global {
  interface Window {
    __PREVIEW_PAYLOAD__?: PreviewPayload;
    setPreviewPayload?: (payload: PreviewPayload) => void;
    pywebview?: { api: PreviewHostApi };
  }
}

export {};
