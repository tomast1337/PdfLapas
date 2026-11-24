// The module PresentationRenderer.renderModelHtml() imports, in a worker thread, to build the toHtml() document
// there (#614); the Node entry names it in toPdf()'s request. Browser-safe, like ./html.
export { modelHtml } from "./render/document";
