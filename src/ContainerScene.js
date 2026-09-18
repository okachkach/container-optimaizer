import { useEffect, useLayoutEffect, useMemo, useRef } from 'react';
import { Canvas, useThree } from '@react-three/fiber';
import { OrbitControls } from '@react-three/drei/core/OrbitControls';
import { BoxGeometry, Color, EdgesGeometry, Object3D } from 'three';
import { productColor } from './services/containerLayout';

function Cartons({ placements }) {
  const mesh = useRef();
  const invalidate = useThree(state => state.invalidate);
  useLayoutEffect(() => {
    if (!mesh.current) return;
    const transform = new Object3D();
    const color = new Color();
    placements.forEach((carton, index) => {
      transform.position.fromArray(carton.position);
      // A small display gap makes adjacent cartons distinguishable; layout sizes retain their full CBM.
      transform.scale.set(...carton.size.map(size => size * 0.985));
      transform.updateMatrix();
      mesh.current.setMatrixAt(index, transform.matrix);
      mesh.current.setColorAt(index, color.set(productColor(carton.productKey)));
    });
    mesh.current.instanceMatrix.needsUpdate = true;
    if (mesh.current.instanceColor) mesh.current.instanceColor.needsUpdate = true;
    mesh.current.computeBoundingSphere();
    invalidate();
  }, [placements, invalidate]);
  if (!placements.length) return null;
  return (
    <instancedMesh key={placements.length} ref={mesh} args={[null, null, placements.length]}>
      <boxGeometry args={[1, 1, 1]} />
      <meshStandardMaterial roughness={0.7} metalness={0.05} />
    </instancedMesh>
  );
}

function Envelope({ size, dark }) {
  const [length, height, width] = size;
  const edges = useMemo(() => {
    const box = new BoxGeometry(length, height, width);
    const geometry = new EdgesGeometry(box);
    box.dispose();
    return geometry;
  }, [length, height, width]);
  // R3F does not dispose primitive objects automatically.
  useEffect(() => () => edges.dispose(), [edges]);
  return (
    <>
      <lineSegments position={[length / 2, height / 2, width / 2]}>
        <primitive object={edges} attach="geometry" />
        <lineBasicMaterial color={dark ? '#94a3b8' : '#64748b'} />
      </lineSegments>
      <mesh position={[length / 2, -height * 0.018, width / 2]}>
        <boxGeometry args={[length, height * 0.035, width]} />
        <meshStandardMaterial color={dark ? '#334155' : '#cbd5e1'} roughness={0.9} />
      </mesh>
    </>
  );
}

export default function ContainerScene({ layout, dark }) {
  const [length, height, width] = layout.containerSize;
  const target = [length / 2, height / 2, width / 2];
  return (
    <Canvas frameloop="demand" dpr={[1, 1.5]} gl={{ antialias: true }}
      camera={{ position: [length * 1.05, length * 0.65, length * 1.25], fov: 42, near: length / 1000, far: length * 30 }}
      fallback={<p role="status" className="p-6 text-center">3D graphics are unavailable in this browser. Switch to List view to continue.</p>}>
      <color attach="background" args={[dark ? '#020617' : '#f8fafc']} />
      <ambientLight intensity={1.2} />
      <directionalLight position={[length / 2, length, length / 2]} intensity={2} />
      <Envelope size={layout.containerSize} dark={dark} />
      <Cartons placements={layout.placements} />
      <OrbitControls makeDefault target={target} minDistance={height * 1.2} maxDistance={length * 4}
        maxPolarAngle={Math.PI / 2} enableDamping dampingFactor={0.12} />
    </Canvas>
  );
}
