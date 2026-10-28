import { cleanup } from "@testing-library/react";
import { afterEach } from "vitest";
import * as antd from "antd";
import * as antdIcons from "@ant-design/icons";
import React from "react";
import ReactDOM from "react-dom";

// jsdom lacks a few browser APIs antd touches during render.
if (!window.matchMedia) {
  window.matchMedia = ((query: string) => ({
    matches: false,
    media: query,
    onchange: null,
    addListener: () => {},
    removeListener: () => {},
    addEventListener: () => {},
    removeEventListener: () => {},
    dispatchEvent: () => false,
  })) as unknown as typeof window.matchMedia;
}

if (!window.HTMLElement.prototype.scrollIntoView) {
  window.HTMLElement.prototype.scrollIntoView = () => {};
}

window.localStorage.setItem("language", "zh");

window.QwenPaw = {
  host: {
    React,
    ReactDOM,
    antd,
    antdIcons,
    apiBaseUrl: "",
    getApiUrl: (path: string) => `/api${path.startsWith("/") ? path : `/${path}`}`,
    getApiToken: () => "test-token",
    getSelectedAgentId: () => "default",
    useTheme: () => "light",
    useLocale: () => "zh",
    fetch: async () => new Response("{}", { status: 200 }),
  },
};

afterEach(() => {
  cleanup();
});
