// Docs site for @komaa/msteams-voice, published to GitHub Pages by .github/workflows/docs.yml.
import { defineConfig } from "astro/config";
import starlight from "@astrojs/starlight";

export default defineConfig({
  site: "https://komaa-com.github.io",
  base: "/openclaw-msteams-voice",
  integrations: [
    starlight({
      title: "MS Teams Voice for OpenClaw",
      description:
        "Microsoft Teams voice and video (CVI) for OpenClaw agents: realtime speech, vision, and a lip-synced avatar, connected through the StandIn media bridge.",
      social: [
        {
          icon: "github",
          label: "GitHub",
          href: "https://github.com/komaa-com/openclaw-msteams-voice",
        },
      ],
      sidebar: [
        { label: "Overview", link: "/" },
        { label: "Getting Started", link: "/getting-started/" },
        { label: "Connecting to StandIn", link: "/connecting-to-standin/" },
        { label: "Architecture", link: "/architecture/" },
        { label: "Configuration Reference", link: "/configuration-reference/" },
        { label: "Realtime and Streaming Modes", link: "/realtime-and-streaming-modes/" },
        { label: "Wire Protocol", link: "/wire-protocol/" },
        { label: "Features", link: "/features/" },
        { label: "Outbound Calls", link: "/outbound-calls/" },
        { label: "Troubleshooting", link: "/troubleshooting/" },
        { label: "Contributing", link: "/contributing/" },
      ],
    }),
  ],
});
