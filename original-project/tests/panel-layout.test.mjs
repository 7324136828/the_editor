import test from "node:test";
import assert from "node:assert/strict";
import { importTypeScript } from "./plugins-test-utils.mjs";

const { DEFAULT_PANEL_LAYOUT, parsePanelLayout, fitPanelLayout } =
  await importTypeScript(
    new URL("../src/services/panelLayout.ts", import.meta.url),
  );
const preferences = (overrides) => ({ ...DEFAULT_PANEL_LAYOUT, ...overrides });

test("missing or damaged saved panel settings recover independent defaults", () => {
  for (const value of [undefined, null, [], "corrupt", 100, false]) {
    const restored = parsePanelLayout(value);
    assert.deepEqual(restored, DEFAULT_PANEL_LAYOUT);
    assert.notEqual(restored, DEFAULT_PANEL_LAYOUT);
  }
  const restored = parsePanelLayout(null);
  restored.primaryWidth = 500;
  assert.equal(
    parsePanelLayout(null).primaryWidth,
    DEFAULT_PANEL_LAYOUT.primaryWidth,
  );
});

test("restoring panel settings rejects non-finite dimensions and invalid visibility flags", () => {
  assert.deepEqual(
    parsePanelLayout({
      primaryWidth: NaN,
      secondaryWidth: Infinity,
      bottomHeight: "400",
      primaryOpen: "false",
      secondaryOpen: 0,
      bottomOpen: null,
    }),
    DEFAULT_PANEL_LAYOUT,
  );
});

test("saved settings preserve explicit visibility choices and clamp out-of-range dimensions", () => {
  const restored = parsePanelLayout({
    primaryWidth: -20,
    secondaryWidth: 100000,
    bottomHeight: 100000,
    primaryOpen: false,
    secondaryOpen: false,
    bottomOpen: true,
  });
  assert.deepEqual(restored, {
    primaryWidth: 100,
    secondaryWidth: 640,
    bottomHeight: 1200,
    primaryOpen: false,
    secondaryOpen: false,
    bottomOpen: true,
  });
  assert.equal(parsePanelLayout({ bottomHeight: -1 }).bottomHeight, 80);
});

test("customized dimensions and visibility round-trip through persisted JSON", () => {
  const saved = preferences({
    primaryWidth: 325,
    secondaryWidth: 410,
    bottomHeight: 320,
    primaryOpen: false,
    bottomOpen: true,
  });
  assert.deepEqual(parsePanelLayout(JSON.parse(JSON.stringify(saved))), saved);
});

test("valid small-window panel sizes survive a save and reload", () => {
  const saved = preferences({
    primaryWidth: 104,
    secondaryWidth: 120,
    bottomHeight: 80,
    secondaryOpen: false,
    bottomOpen: true,
  });
  const before = fitPanelLayout(saved, 224, 240);
  const persisted = {
    ...saved,
    primaryWidth: before.primaryWidth,
    bottomHeight: before.bottomHeight,
  };
  const restored = parsePanelLayout(JSON.parse(JSON.stringify(persisted)));
  const after = fitPanelLayout(restored, 224, 240);
  assert.deepEqual(restored, persisted);
  assert.equal(after.primaryWidth, before.primaryWidth);
  assert.equal(after.bottomHeight, before.bottomHeight);
});

test("a roomy workspace retains preferred sizes and visibility", () => {
  const saved = preferences({
    primaryWidth: 325,
    secondaryWidth: 410,
    bottomHeight: 320,
  });
  const fit = fitPanelLayout(saved, 1600, 900);
  assert.equal(fit.compact, false);
  assert.equal(fit.showPrimary, true);
  assert.equal(fit.showSecondary, true);
  assert.equal(fit.primaryWidth, saved.primaryWidth);
  assert.equal(fit.secondaryWidth, saved.secondaryWidth);
  assert.equal(fit.bottomHeight, saved.bottomHeight);
});

test("narrow workspaces display one sidebar and allow either side to be selected", () => {
  const primary = fitPanelLayout(preferences(), 640, 600);
  assert.equal(primary.compact, true);
  assert.equal(primary.showPrimary, true);
  assert.equal(primary.showSecondary, false);
  const secondary = fitPanelLayout(
    preferences({ primaryOpen: false }),
    640,
    600,
  );
  assert.equal(secondary.showPrimary, false);
  assert.equal(secondary.showSecondary, true);
  const none = fitPanelLayout(
    preferences({ primaryOpen: false, secondaryOpen: false }),
    640,
    600,
  );
  assert.equal(none.showPrimary, false);
  assert.equal(none.showSecondary, false);
});

test("temporary viewport constraints do not mutate saved preferred dimensions", () => {
  const saved = Object.freeze(
    preferences({ primaryWidth: 590, secondaryWidth: 610, bottomHeight: 500 }),
  );
  const small = fitPanelLayout(saved, 900, 600);
  assert.ok(small.primaryWidth < saved.primaryWidth);
  assert.ok(small.secondaryWidth < saved.secondaryWidth);
  assert.ok(small.bottomHeight < saved.bottomHeight);
  const restored = fitPanelLayout(saved, 1920, 1000);
  assert.equal(restored.primaryWidth, 590);
  assert.equal(restored.secondaryWidth, 610);
  assert.equal(restored.bottomHeight, 500);
});

test("a hidden sidebar still has coherent resize bounds when its sibling fills the available dock area", () => {
  for (const [primaryOpen, secondaryOpen] of [
    [false, true],
    [true, false],
  ]) {
    const fit = fitPanelLayout(
      preferences({
        primaryWidth: 640,
        secondaryWidth: 640,
        primaryOpen,
        secondaryOpen,
      }),
      1000,
      600,
    );
    const hidden = primaryOpen ? "secondaryWidth" : "primaryWidth";
    const { min, max } = fit.bounds[hidden];
    assert.ok(min <= max, `${hidden} has inverted bounds ${min}..${max}`);
    assert.ok(
      fit[hidden] >= min && fit[hidden] <= max,
      `${hidden} has invalid fitted width ${fit[hidden]}`,
    );
  }
});

test("all fitted panels use coherent integer bounds at desktop and compact sizes", () => {
  for (const width of [320, 480, 899, 900, 901, 1024, 1280, 1920]) {
    for (const height of [240, 350, 600, 900]) {
      for (const [primaryWidth, secondaryWidth] of [
        [180, 220],
        [260, 285],
        [640, 640],
        [333.5, 492.25],
      ]) {
        for (const [primaryOpen, secondaryOpen] of [
          [true, true],
          [true, false],
          [false, true],
          [false, false],
        ]) {
          const fit = fitPanelLayout(
            preferences({
              primaryWidth,
              secondaryWidth,
              primaryOpen,
              secondaryOpen,
            }),
            width,
            height,
          );
          const details = JSON.stringify({
            width,
            height,
            primaryWidth,
            secondaryWidth,
            primaryOpen,
            secondaryOpen,
            fit,
          });
          for (const key of [
            "primaryWidth",
            "secondaryWidth",
            "bottomHeight",
          ]) {
            const { min, max } = fit.bounds[key];
            assert.ok(
              Number.isInteger(fit[key]) &&
                Number.isInteger(min) &&
                Number.isInteger(max),
              details,
            );
            assert.ok(min <= max, `Inverted ${key} bounds: ${details}`);
            assert.ok(
              fit[key] >= min && fit[key] <= max,
              `${key} lies outside its accessible resize bounds: ${details}`,
            );
          }
          if (!fit.compact) {
            const docks =
              (fit.showPrimary ? fit.primaryWidth + 6 : 0) +
              (fit.showSecondary ? fit.secondaryWidth + 6 : 0);
            assert.ok(
              width - 48 - docks >= 320,
              `Editor must retain 320 pixels: ${details}`,
            );
          } else {
            assert.ok(!(fit.showPrimary && fit.showSecondary), details);
          }
        }
      }
    }
  }
});
