import React from 'react';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import App from './App';
import { readProjects, saveProject, updateHistory } from './projectStorage';

jest.mock('./projectStorage', () => ({ readProjects: jest.fn(), saveProject: jest.fn(), updateHistory: jest.fn() }));

test('restores history after reopening and keeps previous snapshots when opening a copy', async () => {
  Object.defineProperty(global, 'crypto', { configurable: true, value: { randomUUID: () => 'new-snapshot' } });
  const project = { version: 1, projectName: 'Saved shipment', catalog: [], shipment: [], catalogSources: [], containerNames: {},
    capacity: { maxWeight: 27000, maxCbm: 76 }, containers: [{ id: 1, items: [], totalGW: 0, totalNW: 0, totalCBM: 0, totalCartons: 0, totalQty: 0, totalPrice: 0 }] };
  const entry = { id: 'original', name: 'Earlier shipment', createdAt: '2026-01-01T12:00:00Z', project };
  readProjects.mockResolvedValue({ current: project, history: [entry] });
  saveProject.mockResolvedValue();
  updateHistory.mockResolvedValue();
  const view = render(<App />);
  await waitFor(() => expect(screen.getByLabelText('Project name').value).toBe('Saved shipment'));
  fireEvent.click(screen.getByRole('button', { name: /History/ }));
  expect(screen.getByText('Earlier shipment')).toBeTruthy();
  fireEvent.click(screen.getAllByRole('button', { name: 'Open / Duplicate' }).slice(-1)[0]);
  await waitFor(() => expect(screen.getByLabelText('Project name').value).toBe('Earlier shipment'));
  expect(updateHistory).not.toHaveBeenCalled();
  await waitFor(() => expect(saveProject).toHaveBeenCalled());
  expect(saveProject.mock.calls.every(([current]) => current.containers.length === 1)).toBe(true);
  view.unmount();
});
