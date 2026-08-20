# WisWork Office Add-in Lab

This workspace is the development home for WisWork Microsoft Office task-pane add-ins. It follows the small, host-oriented package structure used by [office-agents](https://github.com/hewliyang/office-agents), while keeping the first version focused on a shared Office.js boundary instead of copying its agent runtime.

The development manifest targets Word, Excel, and PowerPoint. The sample task pane detects the current host and can read or replace the current text selection through the shared Office document API.

## Develop

```bash
npm install
npm run dev:office
```

The first run may ask to trust a local development certificate. Keep the HTTPS server running, then sideload [`public/manifest.xml`](public/manifest.xml) into an Office desktop or web host.

## Verify

```bash
npm run test -w @wiswork/office-addin
npm run typecheck -w @wiswork/office-addin
npm run build -w @wiswork/office-addin
```

## Extend

- Keep host-neutral Office.js operations in `src/office-document.ts`.
- Add Word-, Excel-, or PowerPoint-specific adapters in separate files when a feature needs a host-specific requirement set.
- Keep credentials and privileged model calls outside the task pane. An Office renderer must not receive WisWork service keys.
