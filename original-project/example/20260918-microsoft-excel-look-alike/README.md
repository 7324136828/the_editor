# Excel Look-Alike

A Vite/React spreadsheet interface that can run in a browser or as a native Windows desktop app with Electron.

## Development

```powershell
npm install
npm run dev
```

`npm run dev` starts Vite and opens the app in an Electron window. To run only the browser version, use `npm run dev:web`.

## Windows builds

```powershell
# Build and launch the unpackaged desktop app
npm run desktop

# Produce an unpacked Windows application in windows-build/win-unpacked
npm run pack:win

# Produce an NSIS Windows installer in windows-build
npm run dist:win
```
