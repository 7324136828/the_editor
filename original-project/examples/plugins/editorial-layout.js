/**
 * Example trusted JavaScript plugin. Register it on defaultPluginHost in your app
 * bootstrap, or pass it in createPluginHost({ plugins: [...] }).
 * @type {import('../../src/plugins/types').StudioPlugin}
 */
export const editorialLayoutPlugin = {
  id: 'example.editorial',
  name: 'Editorial Layout',
  version: '1.0.0',
  description: 'A readable serif layout for long-form Word documents.',
  permissions: ['document:read', 'document:write'],
  commands: [{
    id: 'example.editorial.apply',
    title: 'Apply editorial layout',
    documentTypes: ['word'],
    run(document) {
      if (document.type !== 'word') throw new Error('This layout requires a Word document.');
      return {
        message: 'Editorial layout applied. Save the document to persist it.',
        document: {
          type: 'word',
          data: {
            ...document.data,
            modifiedAt: new Date().toISOString(),
            typography: { ...document.data.typography, fontFamily: 'Georgia', fontSize: 16, lineHeight: 1.75, marginSize: 'wide' },
            sections: document.data.sections.map(section => ({ ...section, orientation: 'portrait' })),
          },
        },
      };
    },
  }],
};
