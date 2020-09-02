# Node.js Demo

## Regular Node Demo

### Usage

Generate a simple presentation.

```bash
node demo.js
```

Generate a presentation with all demo objects (like the browser demo).

```bash
node demo.js All
```

Generate a presentation with selected demo objects (e.g.: 'Table', 'Text', etc.).  
(See `../common/demos.js` for all tests)

```bash
node demo.js Text
```

## Stream Demo

The `demo_stream.js` file requires the `express` package to demonstrate streaming.

### Usage

```bash
node demo_stream.js
```

Then visit `http://localhost:3000/` on a local web browser to download the streamed file.
