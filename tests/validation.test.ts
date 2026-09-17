import assert from "node:assert/strict";
import test from "node:test";

import { boundedLimit, optionalBoolean, reportMonth, requiredText } from "../src/validation.js";

test("reportMonth accepts only a real first day of month", () => {
  assert.equal(reportMonth("2026-08-01"), "2026-08-01");
  assert.throws(() => reportMonth("2026-08-02"), /first day/);
  assert.throws(() => reportMonth("2026-13-01"), /first day/);
});

test("boundedLimit enforces API limits", () => {
  assert.equal(boundedLimit(undefined), 1000);
  assert.equal(boundedLimit("5000"), 5000);
  assert.throws(() => boundedLimit("0"), /between 1 and 5000/);
  assert.throws(() => boundedLimit("1.5"), /between 1 and 5000/);
});

test("text and boolean payload validation is strict", () => {
  assert.equal(requiredText({ name: "  Cirrus  " }, "name"), "Cirrus");
  assert.throws(() => requiredText({ name: "  " }, "name"), /name is required/);
  assert.equal(optionalBoolean(true, "is_active"), true);
  assert.equal(optionalBoolean(undefined, "is_active"), undefined);
  assert.throws(() => optionalBoolean("true", "is_active"), /must be a boolean/);
});
