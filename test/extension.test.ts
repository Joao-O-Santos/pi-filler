import assert from "node:assert/strict";
import test from "node:test";
import extension from "../src/index.js";

test("extension entry point loads", () => {
	assert.equal(typeof extension, "function");
});
