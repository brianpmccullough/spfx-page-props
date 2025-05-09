# GitHub Copilot Instructions for SPFx and React Development

## General Guidelines
- Use modern JavaScript/TypeScript features (e.g., ES6+).
- Follow the Single Responsibility Principle (SRP) to ensure components and modules have a clear purpose.
- Prefer functional components with React Hooks over class components.
- Use TypeScript for type safety and better developer experience.
- NEVER use `any` for variables and method or function return types.
- Promote creation of services in WebPart.ts or Customizer.ts files, pass services to React components using the services interface.
- Related to point above, avoid use of SharePoint specific concepts in .tsx files to avoid dependencies on SharePoint specifics to promote testability and reusability.
- Minimize use of external packages, but leverage @pnp/sp, @pnp/spfx-controls-react, and @pnp/spfx-property-controls.

## Project Structure
- Organize components into meaningful directories to promote modularity.
- Use a `components` folder for reusable React components.
- Use a `services` folder for API calls or business logic.
- Use a `models` folder for TypeScript interfaces and types.
- Use separate files for each interface, class, type, component, service, etc.

## State Management
- Use React Context for global state management.
- Use React's `useState` and `useReducer` for local component state.

## Styling
- Use CSS Modules or SCSS for component-specific styles.
- Follow BEM (Block Element Modifier) naming conventions for class names.
- Use SPFx themeing for colors and fonts.

## Testing
- Write unit tests for all components using Jest and React Testing Library.
- Mock external dependencies in tests to isolate the unit under test.
- Use snapshot testing for components with static output.

## Accessibility
- Follow WCAG guidelines to ensure accessibility.
- Use semantic HTML elements and ARIA attributes where necessary.
- Test components with screen readers and keyboard navigation.

## Performance
- Use React.memo and `useCallback` to optimize rendering.
- Lazy load components and assets using React's `lazy` and `Suspense`.
- Minimize the use of inline styles and large third-party libraries.

## SPFx Specific Guidelines
- Use the SPFx Yeoman generator to scaffold new projects, webparts, and customizers.
- Use the `@microsoft/sp-http` package for SharePoint API calls.
- Store configuration settings in the `config` folder.
- Use the `@microsoft/sp-core-library` for common SPFx utilities.
- Avoid using `window` or `document` objects.  Get everything you need to run from the SPFx `context`.

## Documentation
- Document all public methods and components using JSDoc or TypeScript comments.
- Maintain a `README.md` file with setup and usage instructions.
- Use inline comments to explain complex logic.

## Version Control
- Use Git for version control and follow GitFlow branching strategy.
- Write meaningful commit messages.
- Use `.gitignore` to exclude unnecessary files from version control.