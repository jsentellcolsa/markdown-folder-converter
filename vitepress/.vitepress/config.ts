import { defineConfig } from 'vitepress';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const docsDir = path.resolve(__dirname, '../../docs');

type SidebarItem = {
  text: string;
  link?: string;
  collapsed?: boolean;
  items?: SidebarItem[];
};

function buildSidebarItems(dir: string, urlBase: string): SidebarItem[] {
  const entries = fs.readdirSync(dir, { withFileTypes: true })
    .filter(e => e.name !== 'sidebar.ts' && e.name !== 'index.md')
    .sort((a, b) => {
      if (a.isDirectory() !== b.isDirectory()) return a.isDirectory() ? -1 : 1;
      return a.name.localeCompare(b.name);
    });

  return entries
    .map((entry): SidebarItem | null => {
      const urlPath = `${urlBase}${entry.name}`;

      if (entry.isDirectory()) {
        return {
          text: entry.name,
          collapsed: true,
          items: buildSidebarItems(path.join(dir, entry.name), urlPath + '/'),
        };
      }

      if (entry.name.endsWith('.md')) {
        return {
          text: entry.name.replace(/\.md$/, ''),
          link: urlPath.replace(/\.md$/, ''),
        };
      }

      return null;
    })
    .filter((item): item is SidebarItem => item !== null);
}

const sidebar = buildSidebarItems(docsDir, '/');

export default defineConfig({
  title: 'COLSA Docs',
  description: 'Converted document library',
  srcDir: '../docs',
  themeConfig: {
    search: {
      provider: 'local',
    },
    sidebar,
    nav: [
      { text: 'Home', link: '/' },
    ],
  },
});
