# Window Management Reference

## Tools

- **`window`**: Control Excel window visibility, position, and state

## Actions

| Action | Purpose | Parameters |
|--------|---------|------------|
| `show` | Make Excel visible and bring to front | *(none)* |
| `hide` | Hide the Excel window | *(none)* |
| `bring-to-front` | Bring Excel to foreground | *(none)* |
| `get-info` | Get window state information | *(none)* |
| `set-state` | Set window state | `windowState` (normal, minimized, maximized) |
| `set-position` | Set position and size | `left`, `top`, `width`, `height` (all optional, in points) |
| `arrange` | Apply preset layout | `preset` (left-half, right-half, top-half, bottom-half, center, full-screen) |
| `set-status-bar` | Show text in Excel status bar | `text` (required — e.g. "Building PivotTable...") |
| `clear-status-bar` | Restore default status bar | *(none)* |

## When to Use Window Management

### Interactive "Agent Mode" — User Watches AI Work in Excel
```
1. window(show)                              → Excel becomes visible
2. window(arrange, preset='right-half')      → Position Excel on right side of screen
3. ... perform Excel operations ...          → User watches changes live
4. window(hide)                              → Hide when done (optional)
```

### Side-by-Side Layout
```
1. window(show)
2. window(arrange, preset='left-half')       → Excel takes left half of screen
   → User's AI assistant occupies the right half
```

### Check Current State
```
1. window(get-info) → Returns visibility, position, size, window state, foreground status
```

## Arrange Presets

| Preset | Position | Use Case |
|--------|----------|----------|
| `left-half` | Left 50% of screen | Side-by-side with AI assistant |
| `right-half` | Right 50% of screen | Side-by-side with AI assistant |
| `top-half` | Top 50% of screen | Stacked view |
| `bottom-half` | Bottom 50% of screen | Stacked view |
| `center` | Centered, 60% of screen | Focused work |
| `full-screen` | Maximized | Full visibility |

## Best Practices

1. **Show before operating visually**: If the user wants to watch operations, call `show` + `arrange` before starting the workflow
2. **Visibility syncs with session**: Show/hide updates session metadata — `file(list)` reflects the current visibility state
3. **Arrange makes visible**: `arrange` automatically shows Excel if it's hidden
4. **set-state makes visible**: Setting state to normal/maximized automatically shows Excel
5. **set-position ensures normal state**: Setting position switches from maximized/minimized to normal automatically
6. **Use get-info to check state**: Before positioning, check if Excel is already visible and where it is

## Common Patterns

### Demo Mode — Show User the Work
```
1. file(open, path='report.xlsx')
2. window(show)
3. window(arrange, preset='left-half')
4. ... create tables, charts, formatting ...
5. file(close, save=true)
   → Excel hidden automatically on close
```

### Quick Peek — Show Result Then Hide
```
1. ... perform operations while hidden ...
2. window(show)                    → Show the result
3. screenshot(capture-sheet)       → Also capture for chat
4. window(hide)                    → Hide again
```

### Status Bar Feedback — Live Progress
```
1. window(show)
2. window(arrange, preset='right-half')
3. window(set-status-bar, text='Writing 500 rows...') → User sees progress
4. range(set-values, ...)
5. window(set-status-bar, text='Building chart...')
6. chart(create-from-range, ...)
7. window(clear-status-bar)                           → Clean up when done
```
