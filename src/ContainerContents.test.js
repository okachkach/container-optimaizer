import React, { useState } from 'react';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom';
import ContainerContents from './ContainerContents';

jest.mock('./ContainerVisualization', () => ({
  __esModule: true,
  default: ({ container }) => <div>Scene for container {container.id}</div>,
}));

function Harness() {
  const [activeId, setActiveId] = useState(null);
  return [1, 2].map(id => (
    <ContainerContents key={id} container={{ id }} maxCbm={76} active={activeId === id} onViewChange={active => setActiveId(active ? id : null)}>
      <input aria-label={`Move quantity ${id}`} defaultValue="1" />
    </ContainerContents>
  ));
}

test('defaults to the list, preserves edited quantities across tabs, and mounts just one scene', async () => {
  render(<Harness />);
  expect(screen.queryByText(/Scene for container/)).not.toBeInTheDocument();
  const quantity = screen.getByLabelText('Move quantity 1');
  fireEvent.change(quantity, { target: { value: '7' } });
  fireEvent.click(screen.getAllByRole('tab', { name: '3D view' })[0]);
  expect(await screen.findByText('Scene for container 1')).toBeVisible();
  expect(quantity).not.toBeVisible();
  fireEvent.click(screen.getAllByRole('tab', { name: '3D view' })[1]);
  await waitFor(() => expect(screen.queryByText('Scene for container 1')).not.toBeInTheDocument());
  expect(screen.getByText('Scene for container 2')).toBeVisible();
  expect(quantity).toBeVisible();
  expect(quantity).toHaveValue('7');
});

test('supports keyboard tab selection and correct panel associations', async () => {
  render(<Harness />);
  const list = screen.getAllByRole('tab', { name: 'List view' })[0];
  fireEvent.keyDown(list, { key: 'ArrowRight' });
  const scene = screen.getAllByRole('tab', { name: '3D view' })[0];
  expect(scene).toHaveFocus();
  expect(scene).toHaveAttribute('aria-selected', 'true');
  expect(await screen.findByText('Scene for container 1')).toBeVisible();
  expect(document.getElementById(scene.getAttribute('aria-controls'))).toHaveAttribute('aria-labelledby', scene.id);
  fireEvent.keyDown(scene, { key: 'Home' });
  expect(list).toHaveFocus();
  expect(list).toHaveAttribute('aria-selected', 'true');
  expect(screen.queryByText('Scene for container 1')).not.toBeInTheDocument();
});
