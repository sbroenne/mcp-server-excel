---
layout: default
title: "Changelog"
description: "Release notes and changelog for ExcelMcp VS Code extension and MCP Server. Track new features, bug fixes, and improvements."
keywords: "Excel MCP changelog, release notes, version history, updates"
permalink: /changelog/
---

<div class="hero">
  <div class="container">
    <div class="hero-content">
      <h1 class="hero-title">Changelog</h1>
      <p class="hero-subtitle">Release notes and version history for Excel MCP Server</p>
    </div>
  </div>
</div>

<div class="container content-section" markdown="1">
{% capture changelog_content %}{% include changelog.md %}{% endcapture %}

{{ changelog_content | markdownify }}
</div>
