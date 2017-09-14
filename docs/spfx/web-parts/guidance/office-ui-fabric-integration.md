---
title: Style your solution with the Office UI Fabric
ms.date: 09/25/2017
ms.prod: sharepoint
---

# Style your solution with the Office UI Fabric

When building solutions on the SharePoint Framework, you should consider styling them using Office UI Fabric. This will allow you to seamlessly integrate your customization with the user experience of SharePoint simplifying end-user adoption.

## Use Office UI Fabric in React

When building SharePoint Framework solutions using React, you can choose whether you want to use React components or only CSS. The advantage of using React components over CSS classes is, that these components contain not only styling but also behavior. Additionally, they encapsulate the correct DOM structure for the given control. Whenever you update to a newer version of Office UI Fabric, you will most likely not need to change your code related to components as any changes to their DOM will be done inside the components.

> **Tip:** When creating projects using the SharePoint Framework Yeoman generator and choosing React as your JavaScript framework, the generator automatically includes a suitable version of the **office-ui-fabric-react** package for you to use in your solution.

### Use Office UI Fabric React components

To use a specific Office UI Fabric React component in your solution, you have to import it first, by specifying its name and path:

```ts
// correct, include only the Button component
import { Button } from 'office-ui-fabric-react/lib/Button';
```

This approach is recommended, as it will include in the generated bundle only the Button component and other objects it requires to work properly, rather than the complete Office UI Fabric React package.

You **should not** reference Office UI Fabric React components like this:

```ts
// incorrect, include all Office UI Fabric
import { Button } from 'office-ui-fabric-react';
```

This approach will result in the whole Office UI Fabric React package being included in the generated bundle. In comparison, when you reference only the Button component, as showed in the first example, the generated bundle will be 204KB. If you use the latter approach and include the whole Office UI Fabric React package, the bundle size will increase to 725KB.

### Use Office UI Fabric React CSS styles

Next to components, Office UI Fabric defines [styles](http://dev.office.com/fabric#/styles) such as layout, typography or icons. There are no React components for using these styles, and to reference them in your solution, you have to reference the corresponding CSS classes.

Office UI Fabric is loaded on the page already by SharePoint so you could simply use the specific CSS classes to use the Fabric styling. By doing this however, you would rely on the version of Office UI Fabric used by SharePoint. If that version would change in a backwards incompatible way, it could lead to breaking changes in your solution.

A more reliable way to use Office UI Fabric in your solution, is by referencing the specific CSS styles and creating your own CSS classes, unique to your solution. It requires more effort than simply referring to existing CSS classes, but it prevents you from risking breaking changes in your customization.

When you create a new project using React with the SharePoint Framework Yeoman generator, it scaffolds the following React component:

```tsx
import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p className="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p className="ms-font-l ms-fontColor-white">{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
```

As you can see, this component relies on Office UI Fabric CSS styles loaded by SharePoint. When the version of Office UI Fabric used by SharePoint changes, it could potentially break the layout of this component.

To avoid this risk, you have to import the required CSS styles from Office UI Fabric and redefine them in classes unique to your solution:

```scss
.helloWorld {
  .container {
    max-width: 700px;
    margin: 0px auto;
    box-shadow: 0 2px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1);
  }

  .row {
    padding: 20px;
  }

  @import "./node_modules/office-ui-fabric/dist/sass/Fabric.Common";

  $grid: 'grid';
  .#{$grid} {
    @include ms-Grid;

    &Row {
      @include ms-Grid-row;
    }

    &Col {
      @include ms-Grid-col;

      .#{$grid} {
        padding: 0;
      }
    }
  }

  @media (min-width: $ms-screen-lg-min) {
    .lg10 {
      @include ms-u-lg10;
    }

    .lgPush1 {
      @include ms-u-lgPush1;
    }
  }

  @media (min-width: $ms-screen-xl-min) {
    .xl8 {
      @include ms-u-xl8;
    }

    .xlPush2 {
      @include ms-u-xlPush2;
    }
  }

  .fontXl {
    @include ms-font-xl;
  }

  .fontL {
    @include ms-font-l;
  }

  .bgColorThemeDark {
    @include ms-bgColor-themeDark;
  }

  .fontColorWhite {
    @include ms-fontColor-white;
  }
}
```

Using the `@import` directive, you can reference all styles used inside Office UI Fabric. Next, you define the specific CSS classes which you use inside your solution. Using `@include` directives, you import the style definitions for those classes from Office UI Fabric.

To use these classes, you change your React component as follows:

```tsx
import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div className={styles.helloWorld}>
        <div className={styles.container}>
          <div className={`${styles.gridRow} ${styles.bgColorThemeDark} ${styles.fontColorWhite} ${styles.row}`}>
            <div className={`${styles.gridCol} ${styles.lg10} ${styles.xl8} ${styles.xlPush2} ${styles.lgPush1}`}>
              <span className={`${styles.fontXl} ${styles.fontColorWhite}`}>Welcome to SharePoint!</span>
              <p className={`${styles.fontL} ${styles.fontColorWhite}`}>Customize SharePoint experiences using Web Parts.</p>
              <p className={`${styles.fontL} ${styles.fontColorWhite}`}>{escape(this.props.description)}</p>
              <PrimaryButton href={'https://aka.ms/spfx'} >Learn more</PrimaryButton>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
```

Notice how string representations of Office UI Fabric classes have been replaced by the strongly typed equivalents from your .scss file. Also, the button has been replaced by the **PrimaryButton** Office UI Fabric React component.

## Use Office UI Fabric in other JavaScript frameworks

React is the only JavaScript framework, for which there is an official package with Office UI Fabric components available. In all other frameworks, you have to use the existing Office UI Fabric design language and build the necessary components yourself. By deconstructing existing React components and looking at the samples provided at [dev.office.com/fabric](http://dev.office.com/fabric), you would build components that fully integrate with your JavaScript framework.

> If you're working with AngularJS, there is the [ngOfficeUIFabric](http://ngofficeuifabric.com) package with Office UI Fabric directives for AngularJS. Because it's based on an older version of Office UI Fabric, you should not use it with the SharePoint Framework as it collides with the version of Office UI Fabric used by SharePoint.

### Don't rely on Office UI Fabric loaded by SharePoint

SharePoint Framework solutions are a part of the page and, theoretically, could reference style sheets loaded by SharePoint and other components. Doing this introduces however a risk. If the referenced style changes, the user experience of the solution could get broken. Changes to CSS styles can be subtle, but can also require a new DOM structure that isn't backwards compatible.

SharePoint Online changes frequently and each change can possibly include a new version of Office UI Fabric that isn't backwards compatible. While SharePoint and its components can anticipate upcoming changes, you cannot and this is why you shouldn't rely upon the version of Office UI Fabric used by SharePoint.

### Don't load Office UI Fabric from CDN

When building add-ins and other standalone applications for Office 365, you might have used Office UI Fabric by loading it from a CDN. You should avoid using this approach when building SharePoint Framework solutions. SharePoint Framework customizations are a part of the page and any style sheets they load from a CDN, apply globally to all elements on the page. By loading Office UI Fabric from a CDN you are at risk of your version colliding with the version used by SharePoint which in worst case will render the whole page useless.
