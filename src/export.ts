import fs from 'fs/promises';
import path from 'path';
import dotenv from 'dotenv';

dotenv.config();

export interface AgentConfig {
  name: string;
  email: string;
  role: string;
  department: string;
  userId: string;
  password: string;
  createdAt: string;
  // Beta-only fields (optional)
  employeeType?: string;
  companyName?: string;
  officeLocation?: string;
  // State tracking fields
  lastAction?: 'CREATE' | 'UPDATE' | 'DELETE' | 'NO_CHANGE';
  lastModified?: string;  // ISO timestamp
  changedFields?: string[];  // Fields that changed in last run (includes custom properties)
  // All standard properties from schema (dynamic)
  [key: string]: any;
  // Custom properties (from open extensions)
  customProperties?: Record<string, any>;
}

export interface AgentsConfigFile {
  agents: AgentConfig[];
  summary: {
    totalAgents: number;
    successfulProvisions: number;
    failedProvisions: number;
    generatedAt: string;
  };
}

export class ConfigExporter {
  private outputDir: string;

  constructor(outputDir: string = 'output') {
    this.outputDir = outputDir;
  }

  /**
   * Export agents configuration to JSON file
   */
  async exportConfig(agents: AgentConfig[], outputPath?: string): Promise<string> {
    const config: AgentsConfigFile = {
      agents,
      summary: {
        totalAgents: agents.length,
        successfulProvisions: agents.length,
        failedProvisions: 0,
        generatedAt: new Date().toISOString(),
      },
    };

    // Ensure output directory exists
    await fs.mkdir(this.outputDir, { recursive: true });

    // Default output path
    const filePath = outputPath || path.join(this.outputDir, 'agents-config.json');

    // Write JSON file with pretty formatting
    await fs.writeFile(filePath, JSON.stringify(config, null, 2), 'utf-8');

    console.log(`✓ Exported configuration to: ${filePath}`);
    console.log(`  Total agents: ${agents.length}`);

    return filePath;
  }

  /**
   * Load existing agents configuration
   */
  async loadConfig(inputPath?: string): Promise<AgentsConfigFile> {
    const filePath = inputPath || path.join(this.outputDir, 'agents-config.json');

    try {
      const content = await fs.readFile(filePath, 'utf-8');
      const config = JSON.parse(content) as AgentsConfigFile;

      console.log(`✓ Loaded configuration from: ${filePath}`);
      console.log(`  Total agents: ${config.agents.length}`);

      return config;
    } catch (error: any) {
      if (error.code === 'ENOENT') {
        throw new Error(`Configuration file not found: ${filePath}`);
      }
      throw error;
    }
  }

  /**
   * Export passwords to separate file (for security)
   */
  async exportPasswords(agents: AgentConfig[], outputPath?: string): Promise<string> {
    const filePath = outputPath || path.join(this.outputDir, 'passwords.txt');

    const lines = [
      '# Agent Passwords',
      '# KEEP THIS FILE SECURE - DO NOT COMMIT TO VERSION CONTROL',
      `# Generated: ${new Date().toISOString()}`,
      '',
      ...agents.map(agent => `${agent.email}\t${agent.password}`),
    ];

    await fs.writeFile(filePath, lines.join('\n'), 'utf-8');

    console.log(`✓ Exported passwords to: ${filePath}`);

    return filePath;
  }


  /**
   * Export to CSV format
   */
  async exportToCsv(agents: AgentConfig[], outputPath?: string): Promise<string> {
    const filePath = outputPath || path.join(this.outputDir, 'agents-export.csv');

    const header = 'name,email,role,department,userId,password,employeeType,companyName,officeLocation,createdAt';
    const rows = agents.map(agent =>
      [
        agent.name,
        agent.email,
        agent.role,
        agent.department,
        agent.userId,
        agent.password,
        agent.employeeType || '',
        agent.companyName || '',
        agent.officeLocation || '',
        agent.createdAt,
      ].join(',')
    );

    const csv = [header, ...rows].join('\n');

    await fs.writeFile(filePath, csv, 'utf-8');

    console.log(`✓ Exported to CSV: ${filePath}`);

    return filePath;
  }

  /**
   * Validate configuration file
   */
  async validateConfig(configPath?: string): Promise<{ valid: boolean; errors: string[] }> {
    const errors: string[] = [];

    try {
      const config = await this.loadConfig(configPath);

      // Check required fields
      if (!config.agents || !Array.isArray(config.agents)) {
        errors.push('Missing or invalid agents array');
      }

      if (!config.summary) {
        errors.push('Missing summary');
      }

      // Validate each agent
      for (let i = 0; i < config.agents.length; i++) {
        const agent = config.agents[i];
        const prefix = `Agent ${i + 1}`;

        if (!agent.name) errors.push(`${prefix}: Missing name`);
        if (!agent.email) errors.push(`${prefix}: Missing email`);
        if (!agent.role) errors.push(`${prefix}: Missing role`);
        if (!agent.department) errors.push(`${prefix}: Missing department`);
        if (!agent.userId) errors.push(`${prefix}: Missing userId`);
        if (!agent.password) errors.push(`${prefix}: Missing password`);
      }

      // Check for duplicate emails
      const emails = config.agents.map(a => a.email);
      const duplicates = emails.filter((email, index) => emails.indexOf(email) !== index);
      if (duplicates.length > 0) {
        errors.push(`Duplicate emails found: ${duplicates.join(', ')}`);
      }

      if (errors.length === 0) {
        console.log('✓ Configuration is valid');
        return { valid: true, errors: [] };
      } else {
        console.error('✗ Configuration validation failed:');
        errors.forEach(error => console.error(`  - ${error}`));
        return { valid: false, errors };
      }
    } catch (error: any) {
      errors.push(`Failed to load config: ${error.message}`);
      return { valid: false, errors };
    }
  }

  /**
   * Generate summary report
   */
  async generateReport(configPath?: string): Promise<string> {
    const config = await this.loadConfig(configPath);

    const lines = [
      '# Agent Provisioning Report',
      `Generated: ${new Date().toISOString()}`,
      '',
      '## Summary',
      `- Total Agents: ${config.agents.length}`,
      `- Successful Provisions: ${config.summary.successfulProvisions}`,
      `- Failed Provisions: ${config.summary.failedProvisions}`,
      `- Provisioned: ${config.summary.generatedAt}`,
      '',
      '## Agents by Department',
    ];

    // Group by department
    const byDepartment = new Map<string, AgentConfig[]>();
    for (const agent of config.agents) {
      if (!byDepartment.has(agent.department)) {
        byDepartment.set(agent.department, []);
      }
      byDepartment.get(agent.department)!.push(agent);
    }

    // Sort departments
    const sortedDepts = Array.from(byDepartment.keys()).sort();

    for (const dept of sortedDepts) {
      const agents = byDepartment.get(dept)!;
      lines.push(`\n### ${dept} (${agents.length})`);
      for (const agent of agents) {
        lines.push(`- ${agent.name} (${agent.role}) - ${agent.email}`);
      }
    }

    const report = lines.join('\n');
    const reportPath = path.join(this.outputDir, 'provisioning-report.md');

    await fs.writeFile(reportPath, report, 'utf-8');

    console.log(`✓ Generated report: ${reportPath}`);

    return reportPath;
  }
}

// CLI support for direct execution
if (import.meta.url === `file://${process.argv[1]}`) {
  const command = process.argv[2];
  const exporter = new ConfigExporter();

  try {
    switch (command) {
      case 'validate':
        await exporter.validateConfig();
        break;

      case 'report':
        await exporter.generateReport();
        break;

      case 'export-csv':
        const config = await exporter.loadConfig();
        await exporter.exportToCsv(config.agents);
        break;

      default:
        console.log('Usage: node export.js [command]');
        console.log('Commands:');
        console.log('  validate      Validate agents-config.json');
        console.log('  report        Generate provisioning report');
        console.log('  export-csv    Export to CSV format');
    }
  } catch (error: any) {
    console.error(`Error: ${error.message}`);
    process.exit(1);
  }
}
