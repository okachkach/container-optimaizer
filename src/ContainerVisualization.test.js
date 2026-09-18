import React from 'react';
import { act, fireEvent, render, screen, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import ContainerVisualization from './ContainerVisualization';

jest.mock('./ContainerScene', () => ({
  __esModule: true,
  default: ({ dark, layout }) => <div data-testid="scene" data-theme={dark ? 'dark' : 'light'}>{layout.placements.length} rendered boxes</div>,
}));

afterEach(() => document.documentElement.classList.remove('dark'));

test('shows actual totals, estimation disclosure, product legend, and unplaced carton accounting', () => {
  const container = { id: 1, totalCBM: 38, items: [
    { markNo: 'A', description: 'Small carton', cbmPerCtn: 0.125 },
    { markNo: 'B', description: 'Unknown carton', cbmPerCtn: 0 },
  ] };
  render(<ContainerVisualization container={container} maxCbm={76} />);
  expect(screen.getByText('50.0%')).toBeVisible();
  expect(screen.getByText('38.0000 / 76.0000 m³')).toBeVisible();
  expect(screen.getByText(/dimensions unavailable/)).toBeVisible();
  expect(screen.getByText('1 cartons unplaced')).toBeVisible();
  expect(screen.getByText('Small carton')).toBeVisible();
  expect(screen.getByText('Unknown carton')).toBeVisible();
});

test('updates an open scene when the theme changes and supports resetting the camera', async () => {
  render(<ContainerVisualization container={{ id: 1, items: [], totalCBM: 0 }} maxCbm={76} />);
  expect(screen.getByTestId('scene')).toHaveAttribute('data-theme', 'light');
  act(() => document.documentElement.classList.add('dark'));
  await waitFor(() => expect(screen.getByTestId('scene')).toHaveAttribute('data-theme', 'dark'));
  const previousScene = screen.getByTestId('scene');
  fireEvent.click(screen.getByRole('button', { name: 'Reset camera' }));
  expect(screen.getByTestId('scene')).not.toBe(previousScene);
  expect(screen.getByText('This container is empty.')).toBeVisible();
});

test('does not mount a scene for invalid capacity', () => {
  render(<ContainerVisualization container={{ id: 1, items: [], totalCBM: 0 }} maxCbm={0} />);
  expect(screen.queryByTestId('scene')).not.toBeInTheDocument();
  expect(screen.getByText(/Set a positive container CBM/)).toBeVisible();
});
