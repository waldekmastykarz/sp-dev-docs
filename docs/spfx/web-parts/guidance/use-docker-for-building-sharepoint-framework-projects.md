# Use Docker for building SharePoint Framework projects

SharePoint projects span from a few weeks to few months depending on their scope. During that period SharePoint Framework can be updated. A newer version of the SharePoint Framework might not be backwards compatible and using the different versions of it in one project can introduce additional migration effort. Using Docker images can help you deal with the different versions of the SharePoint Framework and ensure integrity of all your projects.

## Challenges when working with SharePoint Framework 

SharePoint Framework is a new model for building SharePoint customizations. Unlike other development models available to date it uses build toolchain based on open-source tools. The new framework allows developers on any platform to build SharePoint customizations but it works differently than SharePoint developer tools in the past.  

### Research the SharePoint Framework

SharePoint Framework is currently in developer preview. While many developers are interested in investigating its capabilities, they cannot use it for production development just yet. Instead they build customizations on of the currently supported development models using Visual Studio on Windows.
SharePoint Framework toolchain requires Node.js, NPM and other tools specific to web development, which you might not have installed yet. As you need to be productive, you might not want to install these tools on your primary development machine.

### Investigate the impact of a new version of the framework on an existing project

Migrating an existing project to a new version of the SharePoint Framework might be beneficial for a number of reasons varying from specific bug fixes to using of new capabilities. But updating a project is not without a risk. It could be that the newer version of the SharePoint Framework introduced changes that are not backwards compatible with the existing project and require updating the existing artifacts. One way to verify the impact of migrating a project to a newer version, is to install the latest version of the SharePoint Framework, create a new project and compare it with the existing one. The consequence of this approach is that now you use the latest version of the SharePoint Framework and could potentially break your existing projects.

### Work on new projects and supporting existing projects at the same time

As a developer you often not only work on new projects but also support existing projects. SharePoint Framework projects created in the past might have been created with an older version of the SharePoint Framework incompatible with the latest one. As you can have installed only one version of the SharePoint Framework toolchain it can be difficult to work with different projects built each using a different version of the SharePoint Framework.

If you are a lead developer then it's even more important for you to be able to easily switch between the different versions of the SharePoint Framework used by the particular project.

### Develop Node.js solutions and SharePoint Framework

SharePoint Framework requires Node.js v4 LTS to work. If you're also working on projects based on Node.js you might be using a different version of it. While there are solutions that allow you to run multiple versions of Node.js simultaneously on your developer machine, isolating all other packages along with Node.js is error-prone.

These challenges are not new. In the past already, when building SharePoint solutions, developers used virtual machines configured using the same patch level as the farm of the particular customer, to match the specifications and ensure that solutions they were building would run in production. Using virtual machines to solve the challenges mentioned above is still an option, but there is another, more efficient option.
Because SharePoint Framework works on every platform, Developers can use Docker to isolate their project-specific toolset.

## Use Docker for building SharePoint Framework projects

Docker is a virtualization technology that, just like virtual machines, allows you to run software in an isolated environment. Comparing to virtual machines it's more lightweight: where a typical virtual machine used for building SharePoint add-ins would be at least 20GB, a Docker image for building SharePoint Framework solutions is only around 700MB. Comparing to virtual machines, Docker images are easier to build and distribute amongst the developers which makes them a great tool in the developer's toolbox.

### Docker images and containers

When talking about Docker you regularly hear about **images** and **containers**. Understanding what they are, is essential to using Docker in your development process.

#### Docker images

Docker images are templates with pre-installed software such as Node.js, NPM or Yeoman. Every image is built from a base image, with its own configuration and software. The SharePoint Framework image for example is based on the Node.js 4.6.1 image available on Docker Hub - the public catalog of Docker images that developers can use.

#### Docker files

Docker images are built using Docker files - recipes that contain information about the base image as well as the instructions for Docker about what software should be installed and how the image should be configured. Here is the Docker file used to build the image for the SharePoint Framework.

```
FROM node:4.6.1

MAINTAINER Microsoft

EXPOSE 5432 4321 35729

RUN npm i -g npm && \
    npm i -g gulp yo @microsoft/generator-sharepoint && \
    npm cache clean

RUN useradd --create-home --shell /bin/bash spfx && \
    usermod -aG sudo spfx

USER spfx

ENV HOME /home/spfx
RUN mkdir $HOME/app
WORKDIR $HOME/app
VOLUME $HOME/app

CMD /bin/bash
```

#### Docker containers

To use a Docker image, you start a Docker container using that particular image. Where Docker images are templates, Docker containers are running instances of Docker images. Once you download or build a Docker image, you can start containers off of it in a matter of second without having to download or install any additional software. As long as the Docker container works, you get the access to all software installed in that image, such as Node.js, NPM, Yeoman, gulp, etc.

#### Persisting data

One important detail that you should keep in mind, is that Docker images are read-only. When running Docker containers, whatever you write into the container exists only as long as the container runs. When the container ceases to exist, all changes to the filesystem in the container are discarded as well. When you start a new container off an existing image it won't contain any of your changes done in other containers. If you need to persist data, like create a new SharePoint Framework project, you have to do it by mounting a volume (sharing a folder) from the host and write it into that volume.

### Advantages of using Docker images over virtual machines

Earlier in this article you read about some challenges related to working with the SharePoint Framework. While you could overcome them by using virtual machines, using Docker images instead offers you a number of advantages.

#### Docker images are smaller than virtual machines

Typically a developer virtual machine contains the operating system and all tools necessary for the  development process. For developers building SharePoint add-ins it wasn't exceptional to work with virtual machines of 20GB or more.

Most Docker images are based on Linux and contain only the minimal software required for the particular project. For SharePoint Framework this means a basic Linux OS distribution with Node.js, NPM, Yeoman, gulp and the SharePoint Framework Yeoman generator. The resulting image is just over 700MB only 60MB of which is specific to the SharePoint Framework.

While it is possible to build a SharePoint Framework Docker image based on Windows, there is little benefit of doing it, not to mention that Windows-based images are significantly larger than Linux-based images.

#### Docker containers are more efficient

Starting a virtual machine with a full operating system is slow and often takes a minute or two. In comparison starting a Docker container takes a few seconds at most. Because Docker images are more lightweight than virtual machines, they also require less resources and can be run on less powerful machines than used for SharePoint development in the past.

#### Combine your preferences with standardized toolchain




## TOC

- what is Docker
  Docker is a light-weight virtualization technology that helps you solve a number of problems when working with SharePoint Framework projects.
  - image vs. container
- Docker vs. VMs
  - size: a typical Windows-VM is > 10GB. Linux-based SPFx image is ~700MB
  - size: Docker uses a layered file system (think differencing disks). each layer is tagged. When creating new images underlying layers are downloaded only once. if you'd have multiple images of SPFx only the delta specific to SPFx would be downloaded (~50MB)
  - speed: docker images are small and start in a matter of seconds rather than minutes
  - isolation: each developer can use their own dev env and tools with their specific configuration. When getting a new image, they don't need to reconfigure it to their needs. With Docker they combine their own tools & configuration with the standardized runtime 
- SharePoint Framework challenges solved by Docker
  - research SPFx: no need to install Node.js and other Node-specific tools on your machine
  - developer: when investigating the impact of migrating to a newer version of SPFx install the newer version of the generator sharepoint without altering your mainstream dev environment
  - developer: when supporting projects created in the past use the version of generator sharepoint used to create that project to maintain consistency
  - lead developer: often lead developers work on multiple different projects at the same time. Different projects can use different versions of the generator. given that generators must be installed globally you can have only one version of the specific generator installed at the given time
  - developer: SPFx requires Node.js v4. If you're working with other Node.js projects you might be using a different version. While there are solutions to run multiple versions of node next to each other, using Docker images with a specific node version is the safest
  - developer: you can use the same image to run multiple containers so if you have multiple projects using the same version of SPFx you don't need multiple VMs for each on of them
- getting started with Docker
  - on Windows/macOS
  - running Docker in a VM
  - shared drives to allow the container to create projects on the host
    - despite enabling sharing for the whole drive, the container has only access to the path mapped explicitly through the volume
- running the SharePoint Framework from a Docker container
  - first run (running the container)
    - run docker container
      - image isn't on the machine so it's downloaded from the repository (docker hub by default)
    - $ yo @microsoft/sharepoint
    - $ gulp serve
  - working with the container
    - latest vs. specific version
    - parameters
      - it
      - rm
      - volumes
        - host-path on macOS vs. Windows 
      - ports
      - name
    - what's where
      - commands run in the container
      - everything that's written to the folder mapped as a volume output to the host
      - note: container is based on Linux. Even though node_modules are on the host and you'd have gulp installed globally on the host, it won't work if you're working on an OS other than Linux. Some packages such as Sass have to be compiled for the specific OS, so you'd need to run `npm install` again to compile it on Windows/macOS
  - developing SharePoint Framework solutions using the SharePoint Framework Docker container
    - creating new project
      - create new folder for the project
      - cd to the folder
      - $ code .
      - run the container
      - yo @microsoft/sharepoint
      - gulp serve
        - notice that the browser isn't automatically started because the command has been executed inside the container
      - in the browser go to https://localhost:5432/workbench
      - add the web part to the canvas
      - change something in the code
        - notice how the workbench automatically refreshed the same way you would use it on the host
      - close the container by typing `exit`
        - notice how all project files and node_modules are stored on the host in your project folder
      - commit to source control to share the project with your fellow-developers
    - working with existing project
      - pull from source control
      - cd to the folder
      - code .
      - run the container
      - npm i to restore dependencies
      - gulp serve
      - in the browser go to https://localhost:5432/workbench