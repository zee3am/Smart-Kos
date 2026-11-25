import type { Config } from 'jest';

const config: Config = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/$1',
    '^@generated/prisma$': '<rootDir>/src/generated/prisma/client.ts',
  },
  transform: {
    '^.+\\.tsx?$': [
      'ts-jest',
      {
        tsconfig: {
          jsx: 'react-jsx',
          module: 'commonjs',
          moduleResolution: 'node',
          esModuleInterop: true,
          allowJs: true,
        },
      },
    ],
  },
  collectCoverageFrom: [
    'app/**/actions.ts',
    '!app/actions.ts',
    '!app/dashboard/**/actions.ts',
  ],
  coverageReporters: ['text', 'lcov', 'html'],
};

export default config;
