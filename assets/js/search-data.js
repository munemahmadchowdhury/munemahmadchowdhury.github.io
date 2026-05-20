// get the ninja-keys element
const ninja = document.querySelector('ninja-keys');

// add the home and posts menu items
ninja.data = [{
    id: "nav-about",
    title: "about",
    section: "Navigation",
    handler: () => {
      window.location.href = "/";
    },
  },{id: "nav-blog-amp-code",
          title: "blog &amp; code",
          description: "",
          section: "Navigation",
          handler: () => {
            window.location.href = "/blog/";
          },
        },{id: "nav-publications",
          title: "publications",
          description: "research articles published in peer-reviewed journals.",
          section: "Navigation",
          handler: () => {
            window.location.href = "/publications/";
          },
        },{id: "nav-projects",
          title: "projects",
          description: "Research &amp; consultancy projects.",
          section: "Navigation",
          handler: () => {
            window.location.href = "/projects/";
          },
        },{id: "nav-oped",
          title: "oped",
          description: "Published oped and opinion articles",
          section: "Navigation",
          handler: () => {
            window.location.href = "/oped/";
          },
        },{id: "nav-cv",
          title: "CV",
          description: "Please click on the PDF icon to download the CV.",
          section: "Navigation",
          handler: () => {
            window.location.href = "/cv/";
          },
        },{id: "nav-teaching",
          title: "teaching",
          description: "Courses taught.",
          section: "Navigation",
          handler: () => {
            window.location.href = "/teaching/";
          },
        },{id: "nav-people",
          title: "people",
          description: "People whom I have worked for and with",
          section: "Navigation",
          handler: () => {
            window.location.href = "/people/";
          },
        },{id: "dropdown-bookshelf",
              title: "bookshelf",
              description: "",
              section: "Dropdown",
              handler: () => {
                window.location.href = "/books/";
              },
            },{id: "post-google-gemini-updates-flash-1-5-gemma-2-and-project-astra",
        
          title: 'Google Gemini updates: Flash 1.5, Gemma 2 and Project Astra <svg width="1.2rem" height="1.2rem" top=".5rem" viewBox="0 0 40 40" xmlns="http://www.w3.org/2000/svg"><path d="M17 13.5v6H5v-12h6m3-3h6v6m0-6-9 9" class="icon_svg-stroke" stroke="#999" stroke-width="1.5" fill="none" fill-rule="evenodd" stroke-linecap="round" stroke-linejoin="round"></path></svg>',
        
        description: "We’re sharing updates across our Gemini family of models and a glimpse of Project Astra, our vision for the future of AI assistants.",
        section: "Posts",
        handler: () => {
          
            window.open("https://blog.google/technology/ai/google-gemini-update-flash-ai-assistant-io-2024/", "_blank");
          
        },
      },{id: "post-a-post-with-code-diff",
        
          title: "a post with code diff",
        
        description: "this is how you can display code diffs",
        section: "Posts",
        handler: () => {
          
            window.location.href = "/blog/2024/code-diff/";
          
        },
      },{id: "post-a-post-with-advanced-image-components",
        
          title: "a post with advanced image components",
        
        description: "this is what advanced image components could look like",
        section: "Posts",
        handler: () => {
          
            window.location.href = "/blog/2024/advanced-images/";
          
        },
      },{id: "post-a-post-with-vega-lite",
        
          title: "a post with vega lite",
        
        description: "this is what included vega lite code could look like",
        section: "Posts",
        handler: () => {
          
            window.location.href = "/blog/2024/vega-lite/";
          
        },
      },{id: "post-displaying-external-posts-on-your-al-folio-blog",
        
          title: 'Displaying External Posts on Your al-folio Blog <svg width="1.2rem" height="1.2rem" top=".5rem" viewBox="0 0 40 40" xmlns="http://www.w3.org/2000/svg"><path d="M17 13.5v6H5v-12h6m3-3h6v6m0-6-9 9" class="icon_svg-stroke" stroke="#999" stroke-width="1.5" fill="none" fill-rule="evenodd" stroke-linecap="round" stroke-linejoin="round"></path></svg>',
        
        description: "",
        section: "Posts",
        handler: () => {
          
            window.open("https://medium.com/@al-folio/displaying-external-posts-on-your-al-folio-blog-b60a1d241a0a?source=rss-17feae71c3c4------2", "_blank");
          
        },
      },{id: "books-behavioral-economics",
          title: 'Behavioral Economics',
          description: "",
          section: "Books",handler: () => {
              window.location.href = "/books/behavioral-economics-phillips/";
            },},{id: "news-a-simple-inline-announcement",
          title: 'A simple inline announcement.',
          description: "",
          section: "News",},{id: "news-a-long-announcement-with-details",
          title: 'A long announcement with details',
          description: "",
          section: "News",handler: () => {
              window.location.href = "/news/announcement_2/";
            },},{id: "news-a-simple-inline-announcement-with-markdown-emoji-sparkles-smile",
          title: 'A simple inline announcement with Markdown emoji! :sparkles: :smile:',
          description: "",
          section: "News",},{id: "projects-project-10",
          title: 'project 10',
          description: "Finding the Link Between International Migration and Climate Change- Salinity, Inundation, and River-Erosion Affected Areas in Bangladesh.",
          section: "Projects",handler: () => {
              window.location.href = "/projects/10_project/";
            },},{id: "projects-project-11",
          title: 'project 11',
          description: "Readiness for the Fourth Industrial Revolution (4IR)- Prospects and Challenges of Technical Education of Polytechnic Institutes in Bangladesh",
          section: "Projects",handler: () => {
              window.location.href = "/projects/11_project/";
            },},{id: "projects-project-12",
          title: 'project 12',
          description: "Gender analysis of Waste Value Chain in Chattogram, Cox’s Bazar &amp; Tangail (Plastic Circularity Ecosystem and Circular Economy)",
          section: "Projects",handler: () => {
              window.location.href = "/projects/12_project/";
            },},{id: "projects-project-13",
          title: 'project 13',
          description: "Independent Verification of Enterprise Based Training (EBT) under ASSET Project",
          section: "Projects",handler: () => {
              window.location.href = "/projects/13_project/";
            },},{id: "projects-project-14",
          title: 'project 14',
          description: "Impact Assessment of Capacity Building and Access-to-Finance Interventions for Women Enterpreneurs&#39; RE-Based Livelihood",
          section: "Projects",handler: () => {
              window.location.href = "/projects/14_project/";
            },},{id: "projects-project-1",
          title: 'project 1',
          description: "State of Accountability of State-Owned Enterprises in the Maldives for TI Maldives.",
          section: "Projects",handler: () => {
              window.location.href = "/projects/1_project/";
            },},{id: "projects-project-2",
          title: 'project 2',
          description: "Integrity Assessment of Bangladesh Wash Sector (Water Resources Management, Water Supply, Sanitation, and Climate Change Adaptation).",
          section: "Projects",handler: () => {
              window.location.href = "/projects/2_project/";
            },},{id: "projects-project-3",
          title: 'project 3',
          description: "Gender Transformative Climate Resilient Microfinance Project",
          section: "Projects",handler: () => {
              window.location.href = "/projects/3_project/";
            },},{id: "projects-project-4",
          title: 'project 4',
          description: "Understanding Political and Socio-Economic Impacts due to Climate Change Vulnerability and Explore Potential Capacities of Climate-Vulnerable Communities for Association for Land Reform and Development.",
          section: "Projects",handler: () => {
              window.location.href = "/projects/4_project/";
            },},{id: "projects-project-5",
          title: 'project 5',
          description: "Renewable Energy Finance in Bangladesh- Prospects, Challenges, Risks and De-Risk Mechanisms",
          section: "Projects",handler: () => {
              window.location.href = "/projects/5_project/";
            },},{id: "projects-project-6",
          title: 'project 6',
          description: "Situation Analysis of Indigenous Entrepreneurship in Mymensingh",
          section: "Projects",handler: () => {
              window.location.href = "/projects/6_project/";
            },},{id: "projects-project-7",
          title: 'project 7',
          description: "Baseline Survey of the Community-Based Child Protection Mechanism for the Children of Garment Workers in Bangladesh Project",
          section: "Projects",handler: () => {
              window.location.href = "/projects/7_project/";
            },},{id: "projects-project-8",
          title: 'project 8',
          description: "End-line Evaluation of BRAC’s Gender Responsive Education and Skill Programme",
          section: "Projects",handler: () => {
              window.location.href = "/projects/8_project/";
            },},{id: "projects-project-9",
          title: 'project 9',
          description: "An Assessment and Development Plan Preparation of Playgrounds, Vacant Land, and Greening Open Spaces in Chattogram City Corporation (CCC) Area.",
          section: "Projects",handler: () => {
              window.location.href = "/projects/9_project/";
            },},{id: "teachings-freelance-educator",
          title: 'Freelance Educator',
          description: "Tutored grades 6–12 in English, Math, and Economics, designing lessons and tracking progress to improve exam readiness.",
          section: "Teachings",handler: () => {
              window.location.href = "/teachings/freelance_educator/";
            },},{id: "teachings-teaching-assistant-amp-research-mentor",
          title: 'Teaching Assistant &amp;amp; Research Mentor',
          description: "Gained experience as a Teaching Assistant and Research Mentor while working under Dr. Kazi Mohammed Kamal Uddin as a Graduate Research Assistant.",
          section: "Teachings",handler: () => {
              window.location.href = "/teachings/teaching_assistant/";
            },},{
        id: 'social-cv',
        title: 'CV',
        section: 'Socials',
        handler: () => {
          window.open("/assets/pdf/cv.pdf", "_blank");
        },
      },{
        id: 'social-linkedin',
        title: 'LinkedIn',
        section: 'Socials',
        handler: () => {
          window.open("https://www.linkedin.com/in/munem-ahmad-chowdhury", "_blank");
        },
      },{
        id: 'social-scholar',
        title: 'Google Scholar',
        section: 'Socials',
        handler: () => {
          window.open("https://scholar.google.com/citations?user=A_GZqVMAAAAJ&hl=en", "_blank");
        },
      },{
      id: 'light-theme',
      title: 'Change theme to light',
      description: 'Change the theme of the site to Light',
      section: 'Theme',
      handler: () => {
        setThemeSetting("light");
      },
    },
    {
      id: 'dark-theme',
      title: 'Change theme to dark',
      description: 'Change the theme of the site to Dark',
      section: 'Theme',
      handler: () => {
        setThemeSetting("dark");
      },
    },
    {
      id: 'system-theme',
      title: 'Use system default theme',
      description: 'Change the theme of the site to System Default',
      section: 'Theme',
      handler: () => {
        setThemeSetting("system");
      },
    },];
