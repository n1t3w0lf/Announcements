# Contributing to Announcements Carousel Web Part

Thank you for your interest in contributing to the Announcements Carousel Web Part! This document provides guidelines and information for developers who want to customize, extend, or contribute to this project.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Workflow](#development-workflow)
- [Architecture Overview](#architecture-overview)
- [Coding Standards](#coding-standards)
- [Testing Guidelines](#testing-guidelines)
- [Submitting Changes](#submitting-changes)
- [Common Customizations](#common-customizations)

## Code of Conduct

This project adheres to professional development standards. Contributors are expected to:

- Write clean, maintainable code
- Follow established patterns and conventions
- Document significant changes
- Test thoroughly before submitting
- Be respectful in code reviews and discussions

## Getting Started

### Prerequisites

- Node.js v16.x
- Git
- Visual Studio Code (recommended)
- SPFx development environment
- Basic knowledge of:
  - TypeScript
  - React
  - SharePoint Framework
  - SharePoint REST API

### Setting Up Development Environment

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd Announcements
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start development server**
   ```bash
   gulp serve
   ```

4. **Open in VS Code**
   ```bash
   code .
   ```

### Recommended VS Code Extensions

- ESLint
- TypeScript and JavaScript Language Features
- SPFx Snippets
- SharePoint Framework Snippets

## Development Workflow

### Branch Strategy

- `main`: Production-ready code
- `develop`: Integration branch for features
- `feature/*`: New features
- `bugfix/*`: Bug fixes
- `hotfix/*`: Critical production fixes

### Workflow Steps

1. **Create a feature branch**
   ```bash
   git checkout -b feature/my-new-feature
   ```

2. **Make your changes**
   - Write code
   - Update documentation
   - Add tests

3. **Test locally**
   ```bash
   gulp serve
   ```

4. **Build for production**
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

5. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: Add new celebration icon type"
   ```

6. **Push to repository**
   ```bash
   git push origin feature/my-new-feature
   ```

7. **Create pull request**
   - Provide clear description
   - Link related issues
   - Request reviews

## Architecture Overview

### Directory Structure

```
Announcements/
├── src/
│   └── webparts/
│       └── announcementsCarousel/
│           ├── components/           # React components
│           │   ├── AnnouncementsCarousel.tsx
│           │   ├── AnnouncementsCarousel.module.scss
│           │   └── IAnnouncementsCarouselProps.ts
│           ├── models/               # TypeScript interfaces
│           │   └── IAnnouncement.ts
│           ├── services/             # Business logic
│           │   ├── AnnouncementDataService.ts
│           │   ├── CelebrationIconService.ts
│           │   └── ListProvisioningService.ts
│           ├── loc/                  # Localization
│           │   ├── en-us.js
│           │   └── mystrings.d.ts
│           ├── AnnouncementsCarouselWebPart.ts
│           └── AnnouncementsCarouselWebPart.manifest.json
├── config/                          # SPFx configuration
├── gulpfile.js                      # Build tasks
├── package.json                     # Dependencies
└── tsconfig.json                    # TypeScript config
```

### Key Components

#### 1. Web Part Class
**File**: `AnnouncementsCarouselWebPart.ts`
- Entry point for the web part
- Handles initialization
- Configures property pane
- Renders React component

#### 2. React Component
**File**: `components/AnnouncementsCarousel.tsx`
- Main UI component
- Manages state and carousel logic
- Handles rotation and navigation
- Renders celebration icons

#### 3. Data Service
**File**: `services/AnnouncementDataService.ts`
- SharePoint REST API calls
- CRUD operations
- Data transformation
- Date filtering logic

#### 4. Provisioning Service
**File**: `services/ListProvisioningService.ts`
- Automatic list creation
- Column provisioning
- List configuration
- Verification logic

#### 5. Celebration Icon Service
**File**: `services/CelebrationIconService.ts`
- Icon configuration
- Animation definitions
- Helper methods

### Data Flow

```
User Action
    ↓
Web Part (AnnouncementsCarouselWebPart.ts)
    ↓
React Component (AnnouncementsCarousel.tsx)
    ↓
Data Service (AnnouncementDataService.ts)
    ↓
SharePoint REST API
    ↓
SharePoint List
```

## Coding Standards

### TypeScript Guidelines

1. **Use Strong Typing**
   ```typescript
   // Good
   private loadAnnouncements(): Promise<IAnnouncement[]>

   // Avoid
   private loadAnnouncements(): Promise<any>
   ```

2. **Interfaces Over Type**
   ```typescript
   // Preferred
   export interface IAnnouncement { ... }

   // Less preferred
   export type Announcement = { ... }
   ```

3. **Async/Await Over Promises**
   ```typescript
   // Good
   public async getAnnouncements(): Promise<IAnnouncement[]> {
     const response = await this.context.spHttpClient.get(...);
     return await response.json();
   }

   // Avoid
   public getAnnouncements(): Promise<IAnnouncement[]> {
     return this.context.spHttpClient.get(...)
       .then(response => response.json());
   }
   ```

### React Guidelines

1. **Functional Components with Hooks** (for new components)
   ```typescript
   // Preferred for new components
   const MyComponent: React.FC<IMyProps> = (props) => {
     const [state, setState] = useState<Type>(initialValue);
     // ...
   };
   ```

2. **Class Components** (maintain existing pattern)
   ```typescript
   // Used in current implementation
   export default class AnnouncementsCarousel extends React.Component<IProps, IState> {
     // ...
   }
   ```

3. **PropTypes via TypeScript**
   ```typescript
   export interface IMyComponentProps {
     title: string;
     count?: number; // Optional
   }
   ```

### CSS/SCSS Guidelines

1. **Use SCSS Modules**
   ```scss
   // AnnouncementsCarousel.module.scss
   .announcementsCarousel {
     .imageContainer {
       // Nested styles
     }
   }
   ```

2. **BEM-like Naming**
   ```scss
   .component {}
   .componentElement {}
   .componentElement--modifier {}
   ```

3. **Variables for Repeated Values**
   ```scss
   $primary-color: #0078d4;
   $border-radius: 8px;
   ```

### Naming Conventions

- **Interfaces**: Prefix with `I` (e.g., `IAnnouncement`)
- **Components**: PascalCase (e.g., `AnnouncementsCarousel`)
- **Services**: PascalCase with "Service" suffix (e.g., `DataService`)
- **Functions**: camelCase (e.g., `loadAnnouncements`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `MAX_ITEMS`)
- **Private members**: Prefix with `_` (e.g., `_dataService`)

### Code Organization

1. **Imports Order**
   ```typescript
   // External libraries
   import * as React from 'react';
   import { SPHttpClient } from '@microsoft/sp-http';

   // Internal modules
   import { IAnnouncement } from '../models/IAnnouncement';
   import { DataService } from '../services/DataService';

   // Styles
   import styles from './Component.module.scss';
   ```

2. **Component Structure**
   ```typescript
   export default class MyComponent extends React.Component<IProps, IState> {
     // Static properties
     // Constructor
     // Lifecycle methods
     // Event handlers
     // Private methods
     // Render methods
     // Main render
   }
   ```

## Testing Guidelines

### Unit Testing

(Tests to be added - framework ready)

```typescript
import { AnnouncementDataService } from '../services/AnnouncementDataService';

describe('AnnouncementDataService', () => {
  it('should filter active announcements', async () => {
    // Arrange
    // Act
    // Assert
  });
});
```

### Manual Testing Checklist

Before submitting changes, test:

- [ ] List provisioning on new site
- [ ] List verification on existing site
- [ ] Announcement display with valid dates
- [ ] Announcement hiding with invalid dates
- [ ] Carousel rotation
- [ ] Navigation arrows
- [ ] Navigation dots
- [ ] Play/pause button
- [ ] All celebration icons
- [ ] All icon positions
- [ ] All transition effects
- [ ] Property pane configuration
- [ ] Responsive design (mobile/tablet)
- [ ] Browser compatibility
- [ ] Error handling

## Submitting Changes

### Commit Message Format

Use conventional commits format:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types:**
- `feat`: New feature
- `fix`: Bug fix
- `docs`: Documentation only
- `style`: Code style changes (formatting)
- `refactor`: Code refactoring
- `test`: Adding tests
- `chore`: Maintenance tasks

**Examples:**
```
feat(carousel): Add zoom transition effect

Added new zoom transition option that scales announcements
when transitioning between slides.

Closes #123
```

```
fix(icons): Correct celebration icon positioning

Fixed issue where center-positioned icons were not
properly centered on mobile devices.

Fixes #456
```

### Pull Request Guidelines

1. **Title**: Clear, descriptive title
2. **Description**: Explain what and why
3. **Testing**: Describe testing performed
4. **Screenshots**: Include for UI changes
5. **Breaking Changes**: Clearly document
6. **Documentation**: Update relevant docs

### Pull Request Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
Describe testing performed

## Screenshots (if applicable)
Add screenshots

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No new warnings
- [ ] Tests added/updated
```

## Common Customizations

### Adding a New Celebration Icon

1. **Update the enum** (`models/IAnnouncement.ts`):
   ```typescript
   export enum CelebrationType {
     // ... existing types
     NewYear = 'NewYear'
   }
   ```

2. **Add configuration** (`services/CelebrationIconService.ts`):
   ```typescript
   [CelebrationType.NewYear]: {
     emoji: '🎆',
     backgroundColor: '#FFD700',
     color: '#FFFFFF',
     label: 'New Year',
     animation: 'sparkle'
   }
   ```

3. **Update list provisioning** (`services/ListProvisioningService.ts`):
   ```typescript
   'Choices': {
     'results': [
       // ... existing choices
       'NewYear'
     ]
   }
   ```

### Adding a New Transition Effect

1. **Update props interface** (`components/IAnnouncementsCarouselProps.ts`):
   ```typescript
   transitionEffect: 'fade' | 'slide' | 'zoom' | 'flip';
   ```

2. **Add animation** (`components/AnnouncementsCarousel.module.scss`):
   ```scss
   &.flip {
     animation: flipIn 0.6s ease-in-out;
   }

   @keyframes flipIn {
     from {
       transform: rotateY(90deg);
       opacity: 0;
     }
     to {
       transform: rotateY(0deg);
       opacity: 1;
     }
   }
   ```

3. **Update property pane** (`AnnouncementsCarouselWebPart.ts`):
   ```typescript
   PropertyPaneDropdown('transitionEffect', {
     label: 'Transition Effect',
     options: [
       // ... existing options
       { key: 'flip', text: 'Flip' }
     ]
   })
   ```

### Adding a New Column to the List

1. **Update model** (`models/IAnnouncement.ts`):
   ```typescript
   export interface IAnnouncement {
     // ... existing fields
     Priority: 'High' | 'Medium' | 'Low';
   }
   ```

2. **Add column to provisioning** (`services/ListProvisioningService.ts`):
   ```typescript
   {
     '__metadata': { 'type': 'SP.FieldChoice' },
     'FieldTypeKind': 6,
     'Title': 'Priority',
     'Required': false,
     'Choices': {
       'results': ['High', 'Medium', 'Low']
     },
     'DefaultValue': 'Medium'
   }
   ```

3. **Update data service** (`services/AnnouncementDataService.ts`):
   ```typescript
   // Update select query
   $select=Id,Title,...,Priority

   // Update mapping
   Priority: item.Priority || 'Medium'
   ```

### Adding Localization

1. **Update interface** (`loc/mystrings.d.ts`):
   ```typescript
   declare interface IAnnouncementsCarouselWebPartStrings {
     // ... existing strings
     NewFieldLabel: string;
   }
   ```

2. **Add translations** (`loc/en-us.js`):
   ```javascript
   return {
     // ... existing translations
     "NewFieldLabel": "New Field"
   }
   ```

3. **Use in code**:
   ```typescript
   import * as strings from 'AnnouncementsCarouselWebPartStrings';
   // ...
   label: strings.NewFieldLabel
   ```

## Questions or Issues?

- **Technical Questions**: Create an issue with the "question" label
- **Bug Reports**: Create an issue with the "bug" label
- **Feature Requests**: Create an issue with the "enhancement" label
- **Security Issues**: Contact maintainers directly

## Resources

- [SPFx Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [React Documentation](https://reactjs.org/docs/getting-started.html)
- [TypeScript Documentation](https://www.typescriptlang.org/docs/)
- [Fluent UI Documentation](https://developer.microsoft.com/en-us/fluentui)

---

Thank you for contributing to make this web part better! 🚀
