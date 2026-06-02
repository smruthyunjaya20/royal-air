const fs = require('fs/promises');
const path = require('path');

const rootDir = process.cwd();
const blocksDir = path.join(rootDir, 'blocks');
const modelsDir = path.join(rootDir, 'models');

async function readJson(filePath) {
  const content = await fs.readFile(filePath, 'utf-8');
  return JSON.parse(content);
}

async function writeJson(filePath, data) {
  await fs.writeFile(filePath, `${JSON.stringify(data, null, 2)}\n`, 'utf-8');
}

function mergeById(items) {
  const itemMap = new Map();

  items.forEach((item) => {
    if (item?.id) itemMap.set(item.id, item);
  });

  return [...itemMap.values()];
}

function mergeFilters(filters) {
  const filterMap = new Map();

  filters.forEach((filter) => {
    if (!filter?.id) return;

    if (!filterMap.has(filter.id)) {
      filterMap.set(filter.id, {
        id: filter.id,
        components: [...new Set(filter.components || [])],
      });
      return;
    }

    const existingFilter = filterMap.get(filter.id);
    existingFilter.components = [...new Set([
      ...existingFilter.components,
      ...(filter.components || []),
    ])];
  });

  return [...filterMap.values()];
}

async function getBlockJsonFiles() {
  const entries = await fs.readdir(blocksDir, { withFileTypes: true });
  const files = [];

  await Promise.all(entries.map(async (entry) => {
    if (!entry.isDirectory()) return;

    const blockPath = path.join(blocksDir, entry.name);
    const blockFiles = await fs.readdir(blockPath, { withFileTypes: true });

    blockFiles.forEach((blockFile) => {
      if (blockFile.isFile() && blockFile.name.startsWith('_') && blockFile.name.endsWith('.json')) {
        files.push(path.join(blockPath, blockFile.name));
      }
    });
  }));

  return files.sort();
}

function ensureSectionFilter(baseFilters, blockIds) {
  const mergedFilters = mergeFilters(baseFilters);
  const sectionFilter = mergedFilters.find((filter) => filter.id === 'section');

  if (sectionFilter) {
    sectionFilter.components = [...new Set([
      ...sectionFilter.components,
      ...blockIds,
    ])];
    return mergedFilters;
  }

  return [
    ...mergedFilters,
    {
      id: 'section',
      components: [...new Set(blockIds)],
    },
  ];
}

async function main() {
  const [baseDefinition, baseModels, baseFilters, blockJsonFiles] = await Promise.all([
    readJson(path.join(modelsDir, '_component-definition.json')),
    readJson(path.join(modelsDir, '_component-models.json')),
    readJson(path.join(modelsDir, '_component-filters.json')),
    getBlockJsonFiles(),
  ]);

  const blockJsons = await Promise.all(blockJsonFiles.map(readJson));
  const blockDefinitions = mergeById(blockJsons.flatMap((blockJson) => blockJson.definitions || []));
  const blockModels = mergeById(blockJsons.flatMap((blockJson) => blockJson.models || []));
  const blockFilters = blockJsons.flatMap((blockJson) => blockJson.filters || []);
  const blockIds = blockDefinitions.map((definition) => definition.id);

  const componentDefinition = {
    ...baseDefinition,
    groups: baseDefinition.groups.map((group) => {
      if (group.id !== 'blocks') return group;

      return {
        ...group,
        components: mergeById([
          ...(group.components || []),
          ...blockDefinitions,
        ]),
      };
    }),
  };

  const componentModels = mergeById([
    ...baseModels,
    ...blockModels,
  ]);

  const componentFilters = ensureSectionFilter(
    [
      ...baseFilters,
      ...blockFilters,
    ],
    blockIds,
  );

  await Promise.all([
    writeJson(path.join(rootDir, 'component-definition.json'), componentDefinition),
    writeJson(path.join(rootDir, 'component-models.json'), componentModels),
    writeJson(path.join(rootDir, 'component-filters.json'), componentFilters),
  ]);
}

main().catch((error) => {
  // eslint-disable-next-line no-console
  console.error(error);
  process.exit(1);
});