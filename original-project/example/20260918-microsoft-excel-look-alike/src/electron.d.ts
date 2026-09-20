export {};

declare global {
  interface Window {
    desktopWindow?: {
      minimize: () => void;
      toggleMaximize: () => void;
      close: () => void;
    };
  }
}
