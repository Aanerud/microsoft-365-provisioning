import { describe, it, expect, vi } from 'vitest';
import { loadPropertyDefinitions, getCollectionFieldNames, buildDescriptionsMap, renderCollectionProperties } from '../json-loader.js';

vi.mock('fs', async () => {
  const actual = await vi.importActual<typeof import('fs')>('fs');
  return {
    ...actual,
    readFileSync: vi.fn(),
  };
});

import { readFileSync } from 'fs';

describe('loadPropertyDefinitions', () => {
  it('parses valid JSON array of property definitions', () => {
    const defs = [
      { name: 'Products', type: 'collection', description: 'User products' },
      { name: 'TerritoryTier1', type: 'string', description: 'Region' },
    ];
    vi.mocked(readFileSync).mockReturnValue(JSON.stringify(defs));
    const result = loadPropertyDefinitions('/fake/props.json');
    expect(result).toHaveLength(2);
    expect(result[0].name).toBe('Products');
    expect(result[0].type).toBe('collection');
  });

  it('returns empty array when file does not exist', () => {
    vi.mocked(readFileSync).mockImplementation(() => { throw new Error('ENOENT'); });
    const result = loadPropertyDefinitions('/nonexistent/path.json');
    expect(result).toEqual([]);
  });

  it('returns empty array when JSON is not an array', () => {
    vi.mocked(readFileSync).mockReturnValue('{"not": "array"}');
    const result = loadPropertyDefinitions('/fake/props.json');
    expect(result).toEqual([]);
  });

  it('returns empty array when JSON is malformed', () => {
    vi.mocked(readFileSync).mockReturnValue('not json at all');
    const result = loadPropertyDefinitions('/fake/props.json');
    expect(result).toEqual([]);
  });
});

describe('getCollectionFieldNames', () => {
  it('returns camelCased names of collection-type properties', () => {
    const defs = [
      { name: 'Products', type: 'collection' as const },
      { name: 'TerritoryTier1', type: 'string' as const },
      { name: 'Skills', type: 'collection' as const },
    ];
    const result = getCollectionFieldNames(defs);
    expect(result).toEqual(['products', 'skills']);
  });

  it('returns empty array when no collection properties exist', () => {
    const defs = [
      { name: 'TerritoryTier1', type: 'string' as const },
    ];
    expect(getCollectionFieldNames(defs)).toEqual([]);
  });

  it('handles empty definitions array', () => {
    expect(getCollectionFieldNames([])).toEqual([]);
  });
});

describe('buildDescriptionsMap', () => {
  it('builds camelCase name to description map', () => {
    const defs = [
      { name: 'Products', type: 'collection' as const, description: 'User products' },
      { name: 'TerritoryTier1', type: 'string' as const, description: 'Region level 1' },
    ];
    const map = buildDescriptionsMap(defs);
    expect(map['products']).toBe('User products');
    expect(map['territoryTier1']).toBe('Region level 1');
  });

  it('skips properties without descriptions', () => {
    const defs = [
      { name: 'Products', type: 'collection' as const },
      { name: 'TerritoryTier1', type: 'string' as const, description: 'Region' },
    ];
    const map = buildDescriptionsMap(defs);
    expect(map['products']).toBeUndefined();
    expect(map['territoryTier1']).toBe('Region');
  });
});

describe('renderCollectionProperties', () => {
  it('renders array values as YAML strings', () => {
    const record: any = {
      products: [
        { Name: 'Widget', Model: 'A1' },
        { Name: 'Gadget', Model: 'B2' },
      ],
    };
    renderCollectionProperties(record, ['products']);
    expect(typeof record.products).toBe('string');
    expect(record.products).toContain('- name: Widget');
    expect(record.products).toContain('  model: A1');
    expect(record.products).toContain('- name: Gadget');
  });

  it('sets empty string for empty arrays', () => {
    const record: any = { products: [] };
    renderCollectionProperties(record, ['products']);
    expect(record.products).toBe('');
  });

  it('skips non-array values', () => {
    const record: any = { products: 'not an array' };
    renderCollectionProperties(record, ['products']);
    expect(record.products).toBe('not an array');
  });

  it('handles null record gracefully', () => {
    expect(() => renderCollectionProperties(null, ['products'])).not.toThrow();
  });

  it('filters out null/empty values from YAML output', () => {
    const record: any = {
      products: [{ Name: 'Widget', Model: null, GTIN: '' }],
    };
    renderCollectionProperties(record, ['products']);
    expect(record.products).toContain('name: Widget');
    expect(record.products).not.toContain('model');
    expect(record.products).not.toContain('gtin');
  });
});
