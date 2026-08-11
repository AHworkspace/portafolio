import React, { useEffect, useMemo, useState } from 'react';
import ReactDOM from 'react-dom/client';
import * as XLSX from 'xlsx';
import { isSupabaseEnabled, portfolioId, storageBucket, supabase } from './supabaseClient';
import './styles.css';

type Role = 'visitor' | 'admin';
type SectionId = 'home' | 'profile' | 'tools' | 'projects' | 'contact' | 'settings';
type View = SectionId | 'project-group' | 'project-detail';
type ProjectGroupStyle = 'data' | 'systems' | 'creative';
type ProjectGroupItem = { id: number; name: string; image: string; style?: ProjectGroupStyle };
type GroupWorkItem = { id: number; name: string; image: string };

const tools = ['Python', 'Power BI', 'SQL', 'Figma', 'React', 'Tableau'];
const projectGroups = ['Analisis de Datos y Business Intelligence', 'Sistemas de Informacion', 'Diseno Grafico y Animacion'];
const groupWorks = ['Analisis de ventas y rentabilidad', 'Segmentacion de clientes'];
const navItems: Array<{ view: SectionId; label: string }> = [
  { view: 'home', label: 'Inicio' },
  { view: 'profile', label: 'Perfil' },
  { view: 'tools', label: 'Herramientas' },
  { view: 'projects', label: 'Proyectos' },
  { view: 'contact', label: 'Contacto' },
  { view: 'settings', label: 'Configuracion' },
];

const inferGroupStyle = (group?: Pick<ProjectGroupItem, 'name' | 'style'> | null): ProjectGroupStyle => {
  const name = group?.name.toLowerCase() ?? '';
  if (name.includes('dise') || name.includes('anim') || name.includes('graf') || name.includes('marca') || name.includes('visual') || name.includes('ilustr')) return 'creative';
  if (name.includes('sistema') || name.includes('web') || name.includes('pagina') || name.includes('app')) return 'systems';
  return 'data';
};

const isDataGroupName = (name: string) => {
  const normalized = name.toLowerCase();
  return normalized.includes('analisis de datos') && normalized.includes('business intelligence');
};

const defaultWorksForGroup = (group?: Pick<ProjectGroupItem, 'name' | 'style'> | null): GroupWorkItem[] => {
  const style = inferGroupStyle(group);
  if (style === 'systems') {
    return [
      { id: 1, name: 'Sistema de gestion academica', image: '' },
      { id: 2, name: 'Pagina web institucional', image: '' },
    ];
  }
  if (style === 'creative') {
    return [
      { id: 1, name: 'Identidad visual y piezas graficas', image: '' },
      { id: 2, name: 'Animacion y contenido audiovisual', image: '' },
    ];
  }

  return groupWorks.map((name, index) => ({ id: index + 1, name, image: '' }));
};

function EditMark({ show = true }: { show?: boolean }) {
  return show ? <span className="edit-mark">\</span> : null;
}

function SmallButton({ children, danger = false }: { children: React.ReactNode; danger?: boolean }) {
  return <button className={danger ? 'btn danger' : 'btn'}>{children}</button>;
}

function Login({ onEnter }: { onEnter: (role: Role) => void }) {
  return (
    <main className="login-screen">
      <section className="login-body">
        <div className="login-showcase">
          <form
            className="login-panel"
            onSubmit={(event) => {
              event.preventDefault();
              onEnter('admin');
            }}
          >
            <div className="panel-title">Login</div>
            <label>
              Username
              <input defaultValue="administrador" />
            </label>
            <label>
              Password
              <input type="password" defaultValue="********" />
            </label>
            <div className="remember-row">
              <span>Remember me</span>
              <span>Forgot Password?</span>
            </div>
            <button className="btn primary login-submit" type="submit">
              Login
            </button>
            <p className="register-text">
              Entrar como <button type="button" onClick={() => onEnter('visitor')}>visitante</button>
            </p>
          </form>
        </div>
      </section>
    </main>
  );
}

function Header({
  role,
  active,
  displayMode,
  onToggleDisplayMode,
  onNavigate,
  onLogout,
}: {
  role: Role;
  active: View;
  displayMode: 'day' | 'night';
  onToggleDisplayMode: () => void;
  onNavigate: (view: SectionId) => void;
  onLogout: () => void;
}) {
  const visibleNavItems = role === 'admin'
    ? navItems
    : navItems.filter((item) => item.view !== 'settings');

  return (
    <header className="app-header">
      <nav>
        {visibleNavItems.map((item) => (
          <button
            className={active === item.view ? 'active' : ''}
            key={item.view}
            onClick={() => onNavigate(item.view)}
          >
            {item.label}
          </button>
        ))}
      </nav>
      <div className="session-actions">
        <button onClick={onToggleDisplayMode}>{displayMode === 'night' ? 'modo dia' : 'modo noche'}</button>
        <button>-&gt;</button>
        <button>{role === 'admin' ? 'ADMIN' : 'VISITANTE'}</button>
        <button onClick={onLogout}>cerrar sesion</button>
      </div>
    </header>
  );
}

function Section({
  title,
  children,
  action,
}: {
  title: string;
  children: React.ReactNode;
  action?: React.ReactNode;
}) {
  return (
    <section className="section-box">
      <header>
        <h2>{title}</h2>
      </header>
      <div className="section-content">{children}</div>
      {action && <div className="section-action">{action}</div>}
    </section>
  );
}

function PageBlock({ id, title, children }: { id: SectionId; title: string; children: React.ReactNode }) {
  return (
    <section className="page-block" id={id}>
      <h2 className="page-heading">{title}</h2>
      {children}
    </section>
  );
}

function FlatBox({
  children,
  action,
}: {
  children: React.ReactNode;
  action?: React.ReactNode;
}) {
  return (
    <div className="section-box flat-box">
      <div className="section-content">{children}</div>
      {action && <div className="section-action">{action}</div>}
    </div>
  );
}

function ImagePlaceholder({
  label = '[imagen]',
  admin = true,
  image,
  onImageChange,
}: {
  label?: string;
  admin?: boolean;
  image?: string;
  onImageChange?: (value: string) => void;
}) {
  const uploadImage = (file: File | undefined) => {
    if (!file || !onImageChange) return;
    const reader = new FileReader();
    reader.onload = () => onImageChange(String(reader.result));
    reader.readAsDataURL(file);
  };

  return (
    <div className="image-placeholder">
      {image ? <img src={image} alt={label} /> : <span>{label}</span>}
      {admin && (
        <label className="btn upload-button">
          + Subir imagen
          <input type="file" accept="image/*" onChange={(event) => uploadImage(event.target.files?.[0])} />
        </label>
      )}
    </div>
  );
}

function CoverCarousel({ admin }: { admin: boolean }) {
  const [slides, setSlides] = usePersistentState('portfolio.coverSlides', [
    { id: 1, label: '[imagen de portada 1]', title: 'Portada 1', image: '' },
    { id: 2, label: '[imagen de portada 2]', title: 'Portada 2', image: '' },
    { id: 3, label: '[imagen de portada 3]', title: 'Portada 3', image: '' },
  ]);
  const [activeSlide, setActiveSlide] = useState(0);

  useEffect(() => {
    const timer = window.setInterval(() => {
      setActiveSlide((current) => (current + 1) % slides.length);
    }, 4500);

    return () => window.clearInterval(timer);
  }, [slides.length]);

  return (
    <div className="carousel-shell">
      <div className="carousel-track" style={{ transform: `translateX(-${activeSlide * 100}%)` }}>
        {slides.map((slide) => (
          <ImagePlaceholder
            admin={admin}
            image={slide.image}
            key={slide.id}
            label={slide.label}
            onImageChange={(image) => setSlides((current) => current.map((item) => item.id === slide.id ? { ...item, image } : item))}
          />
        ))}
      </div>
      <div className="carousel-dots" aria-label="Selector de portada">
        {slides.map((slide, index) => (
          <button
            aria-label={slide.title}
            className={index === activeSlide ? 'active' : ''}
            key={slide.title}
            onClick={() => setActiveSlide(index)}
          />
        ))}
      </div>
    </div>
  );
}

function Profile({ admin }: { admin: boolean }) {
  const [name, setName] = usePersistentState('portfolio.profile.name', 'Tu Nombre');
  const [bio, setBio] = usePersistentState(
    'portfolio.profile.bio',
    'Profesional multidisciplinario con experiencia en analisis de datos, desarrollo de sistemas de informacion y diseno UI/UX. Apasionado por transformar datos complejos en soluciones visuales y tecnicas de alto impacto.',
  );
  const [photo, setPhoto] = usePersistentState('portfolio.profile.photo', '');
  const [cv, setCv] = usePersistentState<{ name: string; url: string } | null>('portfolio.profile.cv', null);
  const uploadCv = (file: File | undefined) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setCv({ name: file.name, url: String(reader.result) });
    reader.readAsDataURL(file);
  };

  return (
    <div className="profile-card">
      <div className="avatar">
        {photo ? <img src={photo} alt="Foto de perfil" /> : <span className="avatar-icon">ADM</span>}
        <small>[FOTO]</small>
        {admin && (
          <label className="btn upload-button tiny">
            subir
            <input
              type="file"
              accept="image/*"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = () => setPhoto(String(reader.result));
                reader.readAsDataURL(file);
              }}
            />
          </label>
        )}
      </div>
      <div className="profile-info">
        <h3>
          <EditableText admin={admin} value={name} onChange={setName} />
          <EditMark show={admin} />
        </h3>
        <p>
          <EditableText admin={admin} value={bio} onChange={setBio} multiline />
          <EditMark show={admin} />
        </p>
        <div className="profile-actions">
          {admin && (
            <label className="btn upload-button">
              subir curriculum
              <input
                type="file"
                accept=".pdf,.doc,.docx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                onChange={(event) => uploadCv(event.target.files?.[0])}
              />
            </label>
          )}
          {cv ? (
            <a className="btn" href={cv.url} target="_blank" rel="noreferrer">
              ver curriculum vitae
            </a>
          ) : (
            <button className="btn" disabled>
              ver curriculum vitae
            </button>
          )}
          {cv && <span className="file-name">{cv.name}</span>}
        </div>
      </div>
    </div>
  );
}

function ToolCards({ admin }: { admin: boolean }) {
  const [items, setItems] = usePersistentState('portfolio.tools', [
    { id: 1, name: 'Python', logo: 'PY', image: '' },
    { id: 2, name: 'Power BI', logo: 'BI', image: '' },
    { id: 3, name: 'SQL', logo: 'SQL', image: '' },
    { id: 4, name: 'Figma', logo: 'FIG', image: '' },
    { id: 5, name: 'React', logo: 'RX', image: '' },
    { id: 6, name: 'Tableau', logo: 'TB', image: '' },
  ]);
  const addItem = () => {
    setItems((current) => [...current, { id: Date.now(), name: 'Nueva herramienta', logo: 'NEW', image: '' }]);
  };
  const updateItem = (id: number, patch: Partial<{ name: string; logo: string; image: string }>) => {
    setItems((current) => current.map((item) => item.id === id ? { ...item, ...patch } : item));
  };
  const uploadLogo = (id: number, file: File | undefined) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateItem(id, { image: String(reader.result) });
    reader.readAsDataURL(file);
  };

  return (
    <>
      <div className="tools-grid">
      {items.map((tool) => (
        <article className={admin ? 'tool-card admin-card' : 'tool-card'} key={tool.id}>
          {admin && (
            <div className="tool-admin-actions">
              <label className="btn upload-button tiny">
                logo
                <input type="file" accept="image/*" onChange={(event) => uploadLogo(tool.id, event.target.files?.[0])} />
              </label>
              <button className="mini-delete" onClick={() => setItems((current) => current.filter((item) => item.id !== tool.id))}>x</button>
            </div>
          )}
          <span className="tool-icon">
            {tool.image ? <img src={tool.image} alt={`Logo ${tool.name}`} /> : tool.logo}
          </span>
          <p>
            <EditableText admin={admin} value={tool.name} onChange={(name) => updateItem(tool.id, { name })} />
            <EditMark show={admin} />
          </p>
        </article>
      ))}
      </div>
      {admin && <div className="section-action inline-action"><button className="btn" onClick={addItem}>+ nueva herramienta</button></div>}
    </>
  );
}

function ToolsSection({ admin, boxed = true }: { admin: boolean; boxed?: boolean }) {
  return boxed ? (
    <Section title="Herramientas">
      <ToolCards admin={admin} />
    </Section>
  ) : (
    <FlatBox>
      <ToolCards admin={admin} />
    </FlatBox>
  );
}

function ProjectCards({ admin, onOpenProject }: { admin: boolean; onOpenProject: (group: ProjectGroupItem) => void }) {
  const [items, setItems, groupsReady] = usePersistentState<ProjectGroupItem[]>('portfolio.projectGroups', projectGroups.map((name, index) => ({ id: index + 1, name, image: '', style: inferGroupStyle({ name }) })));
  const [defaultGroupsVersion, setDefaultGroupsVersion, versionReady] = usePersistentState('portfolio.defaultGroups.version', '');
  const normalizedItems = items
    .sort((first, second) => Number(isDataGroupName(second.name)) - Number(isDataGroupName(first.name)));

  useEffect(() => {
    if (!groupsReady || !versionReady || defaultGroupsVersion === 'v3') return;
    const missingGroups = projectGroups
      .filter((name) => !items.some((item) => item.name.toLowerCase() === name.toLowerCase()))
      .map((name, index) => ({ id: Date.now() + index, name, image: '', style: inferGroupStyle({ name }) }));
    if (missingGroups.length > 0) setItems((current) => [...current, ...missingGroups]);
    setDefaultGroupsVersion('v3');
  }, [defaultGroupsVersion, groupsReady, items, setDefaultGroupsVersion, setItems, versionReady]);

  const addGroup = () => {
    const name = askText('Nombre del grupo de proyectos');
    if (name) setItems((current) => [...current, { id: Date.now(), name, image: '' }]);
  };
  const editGroup = (id: number, currentName: string) => {
    const name = askText('Editar grupo de proyectos', currentName);
    if (name) setItems((current) => current.map((item) => item.id === id ? { ...item, name } : item));
  };
  const deleteGroup = (project: ProjectGroupItem) => {
    const confirmed = window.confirm(`Seguro que quieres borrar el grupo "${project.name}"? Esta accion quitara el grupo de la lista.`);
    if (confirmed) setItems((current) => current.filter((item) => item.id !== project.id));
  };

  return (
    <>
      <div className="projects-grid">
      {normalizedItems.map((project, index) => (
        <article className="project-card" key={project.id}>
          <ImagePlaceholder
            admin={admin}
            image={project.image}
            onImageChange={(image) => setItems((current) => current.map((item) => item.id === project.id ? { ...item, image } : item))}
          />
          <div className="project-footer">
            <h3 className={index % 2 === 0 ? 'green' : 'blue'} onClick={() => admin && editGroup(project.id, project.name)}>
              {project.name} <EditMark show={admin} />
            </h3>
            <div className="project-actions">
              <button className="btn" onClick={() => onOpenProject(project)}>
                -&gt; ver
              </button>
              {admin && <button className="btn danger" onClick={() => deleteGroup(project)}>x</button>}
            </div>
          </div>
        </article>
      ))}
      </div>
      {admin && (
        <div className="section-action inline-action">
          <button className="btn" onClick={addGroup}>+ nuevo grupo de proyectos</button>
        </div>
      )}
    </>
  );
}

function ProjectsSection({
  admin,
  onOpenProject,
  compact = false,
  boxed = true,
}: {
  admin: boolean;
  onOpenProject: (group: ProjectGroupItem) => void;
  compact?: boolean;
  boxed?: boolean;
}) {
  return boxed ? (
    <Section title={compact ? 'lista de proyectos' : 'Proyectos'}>
      <ProjectCards admin={admin} onOpenProject={onOpenProject} />
    </Section>
  ) : (
    <FlatBox>
      <ProjectCards admin={admin} onOpenProject={onOpenProject} />
    </FlatBox>
  );
}

function Contact({ admin, boxed = true }: { admin: boolean; boxed?: boolean }) {
  const [links, setLinks] = usePersistentState('portfolio.contactLinks', [
    { id: 1, label: 'email', value: 'tu@email.com' },
    { id: 2, label: 'celular', value: '+1 234 567 890' },
    { id: 3, label: 'linkedin', value: 'linkedin.com/in/tu-perfil' },
    { id: 4, label: 'github', value: 'github.com/tu-usuario' },
  ]);
  const addLink = () => {
    const label = askText('Tipo de contacto');
    if (!label) return;
    const value = askText('Valor del contacto');
    if (!value) return;
    setLinks((current) => [...current, { id: Date.now(), label, value }]);
  };
  const editLink = (id: number, currentLabel: string, currentValue: string) => {
    const label = askText('Editar tipo de contacto', currentLabel);
    if (!label) return;
    const value = askText('Editar valor de contacto', currentValue);
    if (!value) return;
    setLinks((current) => current.map((item) => item.id === id ? { ...item, label, value } : item));
  };
  const body = (
    <>
      <div className="contact-grid">
      {links.map(({ id, label, value }) => (
        <article className="contact-card" key={id}>
          {admin && <button className="mini-delete" onClick={() => setLinks((current) => current.filter((item) => item.id !== id))}>x</button>}
          <span>[{label} \]</span>
          <p onClick={() => admin && editLink(id, label, value)}>
            {value} <EditMark show={admin} />
          </p>
        </article>
      ))}
      </div>
      {admin && <div className="section-action inline-action"><button className="btn" onClick={addLink}>+ nuevo link de contacto</button></div>}
    </>
  );

  return boxed ? <Section title="Contacto">{body}</Section> : <FlatBox>{body}</FlatBox>;
}

function Settings({
  boxed = true,
  theme,
  setTheme,
}: {
  boxed?: boolean;
  theme: PortfolioTheme;
  setTheme: React.Dispatch<React.SetStateAction<PortfolioTheme>>;
}) {
  const updateTheme = (patch: Partial<PortfolioTheme>) => {
    setTheme((current) => ({ ...current, ...patch }));
  };
  const styleControls: Array<{ key: keyof PortfolioTheme; label: string }> = [
    { key: 'background', label: 'Fondo general' },
    { key: 'panel', label: 'Paneles principales' },
    { key: 'panelMuted', label: 'Bloques internos' },
    { key: 'line', label: 'Bordes suaves' },
    { key: 'lineStrong', label: 'Bordes activos' },
    { key: 'ink', label: 'Texto principal' },
    { key: 'muted', label: 'Texto secundario' },
    { key: 'accent', label: 'Acento verde' },
    { key: 'blue', label: 'Acento azul' },
    { key: 'danger', label: 'Eliminar / alerta' },
    { key: 'hoverGlow', label: 'Brillo hover' },
  ];
  const body = (
    <div className="settings-grid">
      <div className="settings-panel">
        <p>SEGURIDAD</p>
        <label>Contrasena actual<input /></label>
        <label>Nueva contrasena<input /></label>
        <label>Confirmar nueva<input /></label>
        <span>OK Contrasena actualizada</span>
        <button className="btn primary">Actualizar contrasena</button>
      </div>
      <div className="settings-panel style-panel">
        <p>ESTILO DEL PORTAFOLIO</p>
        <div className="theme-grid">
          {styleControls.map(({ key, label }) => (
            <label className="color-field" key={key}>
              {label}
              <span className="color-row">
                <input
                  type="color"
                  value={String(theme[key])}
                  onChange={(event) => updateTheme({ [key]: event.target.value } as Partial<PortfolioTheme>)}
                />
                <input
                  value={String(theme[key])}
                  onChange={(event) => updateTheme({ [key]: event.target.value } as Partial<PortfolioTheme>)}
                />
              </span>
            </label>
          ))}
        </div>
        <div className="login-background-control">
          <p>FONDO DEL LOGIN</p>
          <div className="login-background-preview">
            {theme.loginBackground ? <img src={theme.loginBackground.url} alt="Fondo del login" /> : <span>[sin imagen de fondo]</span>}
          </div>
          <div className="preview-actions left-actions">
            <FileUploadButton label="+ subir fondo del login" accept="image/*" onUpload={(asset) => updateTheme({ loginBackground: asset })} />
            {theme.loginBackground && <button className="btn danger" onClick={() => updateTheme({ loginBackground: null })}>quitar fondo</button>}
          </div>
        </div>
        <button className="btn" onClick={() => setTheme(defaultTheme)}>restaurar estilo original</button>
      </div>
      <div className="settings-panel info">
        <p>INFO DEL SISTEMA</p>
        <div><span>Version</span><strong>PORTAFOLIO.SYS v2.5.1</strong></div>
        <div><span>Modo</span><strong>Administrador activo</strong></div>
        <div><span>Fecha</span><strong>2026-08-09</strong></div>
      </div>
    </div>
  );

  return boxed ? <Section title="Configuracion">{body}</Section> : <FlatBox>{body}</FlatBox>;
}

function LandingPage({
  admin,
  onOpenProject,
  theme,
  setTheme,
}: {
  admin: boolean;
  onOpenProject: (group: ProjectGroupItem) => void;
  theme: PortfolioTheme;
  setTheme: React.Dispatch<React.SetStateAction<PortfolioTheme>>;
}) {
  return (
    <>
      <PageBlock id="home" title="Inicio">
        <FlatBox>
          <CoverCarousel admin={admin} />
        </FlatBox>
      </PageBlock>

      <PageBlock id="profile" title="Perfil">
        <Profile admin={admin} />
      </PageBlock>

      <PageBlock id="tools" title="Herramientas">
        <ToolsSection admin={admin} boxed={false} />
      </PageBlock>

      <PageBlock id="projects" title="Proyectos">
        <ProjectsSection admin={admin} onOpenProject={onOpenProject} boxed={false} />
      </PageBlock>

      <PageBlock id="contact" title="Contacto">
        <Contact admin={admin} boxed={false} />
      </PageBlock>

      {admin && (
        <PageBlock id="settings" title="Configuracion">
          <Settings boxed={false} theme={theme} setTheme={setTheme} />
        </PageBlock>
      )}
    </>
  );
}

function ProjectGroup({ admin, group, onOpenProject }: { admin: boolean; group: ProjectGroupItem; onOpenProject: (work: GroupWorkItem) => void }) {
  const groupStyle = inferGroupStyle(group);
  const defaults = defaultWorksForGroup(group);
  const [works, setWorks] = usePersistentState<GroupWorkItem[]>(`portfolio.groupWorks.${group.id}`, defaults);

  useEffect(() => {
    const stillLegacyDefaults = works.length === groupWorks.length && works.every((work, index) => work.name === groupWorks[index] && !work.image);
    if (groupStyle !== 'data' && stillLegacyDefaults) setWorks(defaults);
  }, [defaults, groupStyle, setWorks, works]);

  const restoreAnalysisWorks = () => {
    const missingWorks = groupWorks
      .filter((name) => !works.some((work) => work.name.toLowerCase() === name.toLowerCase()))
      .map((name, index) => ({ id: Date.now() + index, name, image: '' }));
    if (missingWorks.length > 0) setWorks((current) => [...missingWorks, ...current]);
  };

  useEffect(() => {
    if (isDataGroupName(group.name)) restoreAnalysisWorks();
  }, [group.name, works]);

  const addWork = () => {
    const name = askText('Nombre del trabajo');
    if (name) setWorks((current) => [...current, { id: Date.now(), name, image: '' }]);
  };
  const editWork = (index: number, currentName: string) => {
    const name = askText('Editar trabajo', currentName);
    if (name) setWorks((current) => current.map((work, workIndex) => workIndex === index ? { ...work, name } : work));
  };
  const updateWorkImage = (index: number, image: string) => {
    setWorks((current) => current.map((item, workIndex) => workIndex === index ? { ...item, image } : item));
  };
  const deleteWork = (index: number) => {
    setWorks((current) => current.filter((_, workIndex) => workIndex !== index));
  };

  return (
    <div className={`group-view group-view-${groupStyle}`}>
      <h1 className="title-line">
        {group.name} <EditMark show={admin} />
      </h1>
      {groupStyle === 'systems' && (
        <div className="systems-group-intro">
          <span>Sistemas / paginas web</span>
          <p>Trabajos preparados para mostrar demo, pantallas, modulos y pasos de prueba.</p>
        </div>
      )}
      {groupStyle === 'creative' && (
        <div className="creative-group-intro">
          <span>Diseno grafico / animacion</span>
          <p>Trabajos pensados para mostrar piezas visuales, identidad, proceso creativo y entregables.</p>
        </div>
      )}
      <Section
        title="lista de proyectos"
        action={admin ? (
          <>
            {isDataGroupName(group.name) && <button className="btn" onClick={restoreAnalysisWorks}>restaurar trabajos base</button>}
            <button className="btn" onClick={addWork}>+ anadir nuevo proyecto</button>
          </>
        ) : undefined}
      >
        <div className="work-grid">
          {works.map((work, index) => (
            <article className="work-card" key={`${work.id}-${index}`}>
              <ImagePlaceholder
                admin={admin}
                image={work.image}
                onImageChange={(image) => updateWorkImage(index, image)}
              />
              <div className="work-footer">
                <h3 onClick={() => admin && editWork(index, work.name)}>
                  {work.name} <EditMark show={admin} />
                </h3>
                <div className="work-actions">
                  <button className="btn enter-project" onClick={() => onOpenProject(work)}>
                    entrar al proyecto
                  </button>
                  {admin && <button className="btn danger" onClick={() => deleteWork(index)}>x</button>}
                </div>
              </div>
            </article>
          ))}
        </div>
      </Section>
    </div>
  );
}

type DetailSectionKey = 'tools' | 'presentation' | 'resources' | 'processes' | 'architecture' | 'queries';

type DetailSectionConfig = {
  id: number;
  type: DetailSectionKey;
  title: string;
  visible: boolean;
};

type DetailTool = {
  id: number;
  name: string;
  description: string;
  logo: string;
  image?: string;
};

type FileAsset = {
  id: number;
  name: string;
  type: string;
  url: string;
};

type StoredFile = FileAsset | string;

type ArchitecturePreview = {
  id: number;
  sourceKey: string;
  file: StoredFile;
};

type StoredArchitecturePreview = ArchitecturePreview | StoredFile;

type ProcessItem = {
  id: number;
  title: string;
  description: string;
  files: StoredFile[];
};

type ArchitectureModel = {
  id: number;
  title: string;
  segments: ProcessItem[];
};

type QueryItem = {
  id: number;
  question: string;
  answer: string;
  preview?: StoredFile | null;
  previews?: StoredArchitecturePreview[];
};

type ViewerContext = {
  file: StoredFile;
  previewSourceKey?: string;
};

type PortfolioTheme = {
  background: string;
  panel: string;
  panelMuted: string;
  line: string;
  lineStrong: string;
  ink: string;
  muted: string;
  accent: string;
  blue: string;
  danger: string;
  hoverGlow: string;
  loginBackground: FileAsset | null;
};

const askText = (message: string, current = '') => {
  const value = window.prompt(message, current);
  return value === null ? null : value.trim();
};

const fileName = (file: StoredFile) => typeof file === 'string' ? file : file.name;
const fileKey = (file: StoredFile) => typeof file === 'string' ? file : `file-${file.id}`;
const architectureSourceKey = (modelId: number, segmentId: number, file: StoredFile) => `model-${modelId}-segment-${segmentId}-${fileKey(file)}`;
const processSourceKey = (processId: number, file: StoredFile) => `process-${processId}-${fileKey(file)}`;
const matchesSourceKey = (sourceKey: string, file: StoredFile) => (
  sourceKey === fileKey(file) || (typeof file !== 'string' && sourceKey.startsWith(`${file.id}-`))
);
const matchesArchitectureSourceKey = (sourceKey: string, modelId: number, segmentId: number, file: StoredFile) => (
  sourceKey === architectureSourceKey(modelId, segmentId, file) || matchesSourceKey(sourceKey, file)
);
const sameStoredFile = (first: StoredFile, second: StoredFile) => {
  if (typeof first === 'string' || typeof second === 'string') return first === second;
  return first.id === second.id;
};
const isArchitecturePreview = (preview: StoredArchitecturePreview): preview is ArchitecturePreview => (
  typeof preview === 'object' && 'sourceKey' in preview && 'file' in preview
);
const isImageFile = (file: StoredFile | null | undefined) => typeof file !== 'string' && Boolean(file?.type.startsWith('image/'));
const fileExtension = (file: StoredFile | null | undefined) => file ? fileName(file).split('.').pop()?.toLowerCase() ?? '' : '';
const isPdfFile = (file: StoredFile | null | undefined) => {
  if (!file || typeof file === 'string') return false;
  return file.type === 'application/pdf' || fileExtension(file) === 'pdf';
};
const isTextFile = (file: StoredFile | null | undefined) => {
  if (!file || typeof file === 'string') return false;
  const extension = fileExtension(file);
  return file.type.startsWith('text/') || ['csv', 'txt', 'md', 'json'].includes(extension);
};
const isExcelFile = (file: StoredFile | null | undefined) => {
  if (!file || typeof file === 'string') return false;
  const extension = fileExtension(file);
  return ['xls', 'xlsx', 'xlsm', 'csv'].includes(extension);
};
const isPowerBiProject = (file: StoredFile | null | undefined) => fileExtension(file) === 'pbix';
const isFileAsset = (value: unknown): value is FileAsset => (
  value !== null
  && typeof value === 'object'
  && 'id' in value
  && 'name' in value
  && 'type' in value
  && 'url' in value
);

const openFileDb = () => new Promise<IDBDatabase>((resolve, reject) => {
  const request = indexedDB.open('portfolio-files', 1);
  request.onupgradeneeded = () => {
    request.result.createObjectStore('files');
  };
  request.onsuccess = () => resolve(request.result);
  request.onerror = () => reject(request.error);
});

const saveFileAsset = async (asset: FileAsset) => {
  const db = await openFileDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction('files', 'readwrite');
    tx.objectStore('files').put(asset.url, asset.id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
};

const loadRemoteState = async <T,>(key: string): Promise<T | null> => {
  if (!supabase) return null;

  const { data, error } = await supabase
    .from('portfolio_state')
    .select('value')
    .eq('portfolio_id', portfolioId)
    .eq('key', key)
    .maybeSingle();

  if (error) {
    console.warn('No se pudo cargar estado desde Supabase:', error.message);
    return null;
  }

  return data?.value as T ?? null;
};

const saveRemoteState = async (key: string, value: unknown) => {
  if (!supabase) return;

  const { error } = await supabase
    .from('portfolio_state')
    .upsert({
      portfolio_id: portfolioId,
      key,
      value,
      updated_at: new Date().toISOString(),
    }, { onConflict: 'portfolio_id,key' });

  if (error) console.warn('No se pudo guardar estado en Supabase:', error.message);
};

const dataUrlToBlob = async (dataUrl: string) => {
  const response = await fetch(dataUrl);
  return response.blob();
};

const saveSupabaseFileAsset = async (asset: FileAsset): Promise<FileAsset> => {
  if (!supabase || !asset.url.startsWith('data:')) return asset;

  const safeName = asset.name.replace(/[^\w.-]+/g, '-');
  const path = `${portfolioId}/${asset.id}-${safeName}`;
  const blob = await dataUrlToBlob(asset.url);
  const { error } = await supabase.storage
    .from(storageBucket)
    .upload(path, blob, {
      contentType: asset.type,
      upsert: true,
    });

  if (error) {
    console.warn('No se pudo subir archivo a Supabase Storage:', error.message);
    return asset;
  }

  const { data } = supabase.storage.from(storageBucket).getPublicUrl(path);
  return { ...asset, url: data.publicUrl };
};

const loadFileAssetUrl = async (id: number) => {
  const db = await openFileDb();
  const url = await new Promise<string>((resolve, reject) => {
    const request = db.transaction('files', 'readonly').objectStore('files').get(id);
    request.onsuccess = () => resolve(String(request.result || ''));
    request.onerror = () => reject(request.error);
  });
  db.close();
  return url;
};

const sanitizePersistentValue = (value: unknown) => JSON.stringify(value, (_key, current) => {
  if (isFileAsset(current)) return { ...current, url: '' };
  return current;
});

const hydrateFileAssets = async <T,>(value: T): Promise<T> => {
  let changed = false;

  const hydrate = async (current: unknown): Promise<unknown> => {
    if (isFileAsset(current) && !current.url) {
      const url = await loadFileAssetUrl(current.id);
      if (url) {
        changed = true;
        return { ...current, url };
      }
      return current;
    }

    if (Array.isArray(current)) {
      return Promise.all(current.map((item) => hydrate(item)));
    }

    if (current && typeof current === 'object') {
      const entries = await Promise.all(Object.entries(current).map(async ([key, item]) => [key, await hydrate(item)]));
      return Object.fromEntries(entries);
    }

    return current;
  };

  const hydrated = await hydrate(value) as T;
  return changed ? hydrated : value;
};

const readFileAsset = (file: File, onDone: (asset: FileAsset) => void) => {
  const reader = new FileReader();
  reader.onload = async () => {
    const asset = {
      id: Date.now(),
      name: file.name,
      type: file.type || 'application/octet-stream',
      url: String(reader.result),
    };
    const storedAsset = isSupabaseEnabled ? await saveSupabaseFileAsset(asset) : asset;
    saveFileAsset(storedAsset).finally(() => onDone(storedAsset));
  };
  reader.readAsDataURL(file);
};

const renameStoredFile = (file: StoredFile, name: string): StoredFile => (
  typeof file === 'string' ? name : { ...file, name }
);

const defaultTheme: PortfolioTheme = {
  background: '#edf3fa',
  panel: '#fbfdff',
  panelMuted: '#f2f6fc',
  line: '#b7d0ea',
  lineStrong: '#8fb9de',
  ink: '#001733',
  muted: '#5e819d',
  accent: '#00c987',
  blue: '#008fe8',
  danger: '#e23d3d',
  hoverGlow: '#00ffaa',
  loginBackground: null,
};

const hexToRgb = (hex: string) => {
  const clean = hex.replace('#', '');
  const normalized = clean.length === 3 ? clean.split('').map((char) => char + char).join('') : clean;
  const value = Number.parseInt(normalized, 16);
  if (Number.isNaN(value)) return '251, 253, 255';
  return `${(value >> 16) & 255}, ${(value >> 8) & 255}, ${value & 255}`;
};

const normalizeTheme = (theme: Partial<PortfolioTheme>): PortfolioTheme => ({
  ...defaultTheme,
  ...theme,
  loginBackground: theme.loginBackground ?? null,
});

const sectionTypeLabels: Record<DetailSectionKey, string> = {
  tools: 'Herramientas',
  presentation: 'Presentacion',
  resources: 'Recursos',
  processes: 'Limpieza y transformacion',
  architecture: 'Arquitectura del proyecto',
  queries: 'Consultas',
};

const sectionTypes: DetailSectionKey[] = ['tools', 'presentation', 'resources', 'processes', 'architecture', 'queries'];

const getInitialSectionLayout = (): DetailSectionConfig[] => {
  try {
    const savedLayout = window.localStorage.getItem('portfolio.detail.sectionLayout');
    if (savedLayout) return JSON.parse(savedLayout) as DetailSectionConfig[];

    const legacySections = window.localStorage.getItem('portfolio.detail.sections');
    const parsedLegacy = legacySections
      ? JSON.parse(legacySections) as Record<DetailSectionKey, { title: string; visible: boolean }>
      : null;

    return sectionTypes.map((type, index) => ({
      id: index + 1,
      type,
      title: parsedLegacy?.[type]?.title ?? sectionTypeLabels[type],
      visible: parsedLegacy?.[type]?.visible ?? true,
    }));
  } catch {
    return sectionTypes.map((type, index) => ({ id: index + 1, type, title: sectionTypeLabels[type], visible: true }));
  }
};

const defaultArchitectureSegment: ProcessItem = {
  id: 1,
  title: 'Segmentacion',
  description: 'Agrupacion de clientes k=4',
  files: ['BD POWER BI PRACTICAS.xlsx'],
};

const getInitialArchitectureModels = (): ArchitectureModel[] => {
  try {
    const savedModels = window.localStorage.getItem('portfolio.detail.architectureModels');
    if (savedModels) return JSON.parse(savedModels) as ArchitectureModel[];

    const legacyModelTitle = window.localStorage.getItem('portfolio.detail.modelTitle');
    const legacyProcess = window.localStorage.getItem('portfolio.detail.architectureProcess');

    return [{
      id: 1,
      title: legacyModelTitle ? JSON.parse(legacyModelTitle) as string : 'Modelo K-Means',
      segments: [legacyProcess ? JSON.parse(legacyProcess) as ProcessItem : defaultArchitectureSegment],
    }];
  } catch {
    return [{ id: 1, title: 'Modelo K-Means', segments: [defaultArchitectureSegment] }];
  }
};

function usePersistentState<T>(key: string, initialValue: T) {
  const [value, setValue] = useState<T>(() => {
    try {
      const saved = window.localStorage.getItem(key);
      return saved ? JSON.parse(saved) as T : initialValue;
    } catch {
      return initialValue;
    }
  });
  const [remoteReady, setRemoteReady] = useState(!isSupabaseEnabled);

  useEffect(() => {
    let cancelled = false;

    if (!isSupabaseEnabled) return;

    loadRemoteState<T>(key)
      .then((remoteValue) => {
        if (!cancelled && remoteValue !== null) setValue(remoteValue);
      })
      .finally(() => {
        if (!cancelled) setRemoteReady(true);
      });

    return () => {
      cancelled = true;
    };
  }, [key]);

  useEffect(() => {
    try {
      window.localStorage.setItem(key, sanitizePersistentValue(value));
    } catch {
      // Large uploads can exceed browser storage. The UI still works for the current session.
    }
  }, [key, value]);

  useEffect(() => {
    if (!remoteReady) return;
    saveRemoteState(key, value);
  }, [key, remoteReady, value]);

  return [value, setValue, remoteReady] as const;
}

function useHydratedPersistentState<T>(key: string, initialValue: T) {
  const [value, setValue] = usePersistentState<T>(key, initialValue);

  useEffect(() => {
    let cancelled = false;
    hydrateFileAssets(value).then((hydrated) => {
      if (!cancelled && hydrated !== value) setValue(hydrated);
    });

    return () => {
      cancelled = true;
    };
  }, [value, setValue]);

  return [value, setValue] as const;
}

function EditableText({
  admin,
  value,
  onChange,
  className,
  multiline = false,
}: {
  admin: boolean;
  value: string;
  onChange: (value: string) => void;
  className?: string;
  multiline?: boolean;
}) {
  if (!admin) {
    return <span className={className}>{value}</span>;
  }

  return multiline ? (
    <textarea
      className={className ? `inline-field ${className}` : 'inline-field'}
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  ) : (
    <input
      className={className ? `inline-field ${className}` : 'inline-field'}
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  );
}

type SystemModule = {
  id: number;
  title: string;
  description: string;
};

type SystemScreen = {
  id: number;
  title: string;
  description: string;
  image: string;
};

type SystemStep = {
  id: number;
  text: string;
};

type CreativePiece = {
  id: number;
  title: string;
  description: string;
  image: string;
};

type CreativeColor = {
  id: number;
  name: string;
  value: string;
};

type CreativeProcessStep = {
  id: number;
  title: string;
  description: string;
};

function SystemsProjectDetail({ admin, group, work }: { admin: boolean; group: ProjectGroupItem | null; work: GroupWorkItem | null }) {
  const scope = `${group?.id ?? 'systems'}-${work?.id ?? 'main'}`;
  const [title, setTitle] = usePersistentState(`portfolio.systems.${scope}.title`, work?.name ?? 'Sistema de gestion academica');
  const [summary, setSummary] = usePersistentState(
    `portfolio.systems.${scope}.summary`,
    'Sistema web preparado para mostrar objetivos, flujo de uso, pantallas principales y una demo facil de probar por el visitante.',
  );
  const [demoUrl, setDemoUrl] = usePersistentState(`portfolio.systems.${scope}.demoUrl`, 'https://tu-demo.com');
  const [demoUser, setDemoUser] = usePersistentState(`portfolio.systems.${scope}.demoUser`, 'demo@usuario.com');
  const [demoPassword, setDemoPassword] = usePersistentState(`portfolio.systems.${scope}.demoPassword`, '123456');
  const [modules, setModules] = usePersistentState<SystemModule[]>(`portfolio.systems.${scope}.modules`, [
    { id: 1, title: 'Panel principal', description: 'Vista de resumen con indicadores y accesos rapidos.' },
    { id: 2, title: 'Gestion de registros', description: 'Alta, edicion, busqueda y eliminacion de informacion.' },
    { id: 3, title: 'Reportes', description: 'Lectura clara de resultados, estados y exportaciones.' },
  ]);
  const [screens, setScreens] = usePersistentState<SystemScreen[]>(`portfolio.systems.${scope}.screens`, [
    { id: 1, title: 'Inicio del sistema', description: 'Primera pantalla del flujo principal.', image: '' },
    { id: 2, title: 'Modulo de gestion', description: 'Pantalla donde el usuario realiza las acciones clave.', image: '' },
  ]);
  const [steps, setSteps] = usePersistentState<SystemStep[]>(`portfolio.systems.${scope}.steps`, [
    { id: 1, text: 'Abrir la demo y entrar con las credenciales indicadas.' },
    { id: 2, text: 'Revisar el panel principal y navegar por los modulos.' },
    { id: 3, text: 'Crear un registro de prueba y verificar que se muestre correctamente.' },
  ]);

  const addModule = () => setModules((current) => [...current, { id: Date.now(), title: 'Nuevo modulo', description: 'Describe que hace este modulo.' }]);
  const addScreen = () => setScreens((current) => [...current, { id: Date.now(), title: 'Nueva pantalla', description: 'Describe que se puede ver o probar aqui.', image: '' }]);
  const addStep = () => setSteps((current) => [...current, { id: Date.now(), text: 'Nuevo paso de prueba.' }]);

  return (
    <div className="group-view group-view-systems">
      <p className="breadcrumb">Proyectos / {group?.name ?? 'Sistemas de Informacion'}</p>
      <h1 className="title-line">
        <EditableText admin={admin} value={title} onChange={setTitle} />
        <EditMark show={admin} />
      </h1>

      <section className="systems-hero-panel">
        <div>
          <span className="systems-kicker">Proyecto web funcional</span>
          <p>
            <EditableText admin={admin} value={summary} onChange={setSummary} multiline />
            <EditMark show={admin} />
          </p>
        </div>
        <div className="systems-demo-card">
          <span>DEMO</span>
          <EditableText admin={admin} value={demoUrl} onChange={setDemoUrl} />
          <a className="btn primary" href={demoUrl} target="_blank" rel="noreferrer">abrir demo</a>
        </div>
      </section>

      <Section title="Acceso de prueba">
        <div className="systems-demo-grid">
          <article className="systems-credential">
            <span>USUARIO</span>
            <p><EditableText admin={admin} value={demoUser} onChange={setDemoUser} /></p>
          </article>
          <article className="systems-credential">
            <span>CONTRASENA</span>
            <p><EditableText admin={admin} value={demoPassword} onChange={setDemoPassword} /></p>
          </article>
        </div>
      </Section>

      <Section title="Modulos principales" action={admin ? <button className="btn" onClick={addModule}>+ anadir modulo</button> : undefined}>
        <div className="systems-module-grid">
          {modules.map((module) => (
            <article className="systems-module-card" key={module.id}>
              {admin && <button className="mini-delete" onClick={() => setModules((current) => current.filter((item) => item.id !== module.id))}>x</button>}
              <h3>
                <EditableText
                  admin={admin}
                  value={module.title}
                  onChange={(title) => setModules((current) => current.map((item) => item.id === module.id ? { ...item, title } : item))}
                />
              </h3>
              <p>
                <EditableText
                  admin={admin}
                  value={module.description}
                  onChange={(description) => setModules((current) => current.map((item) => item.id === module.id ? { ...item, description } : item))}
                  multiline
                />
              </p>
            </article>
          ))}
        </div>
      </Section>

      <Section title="Pantallas del sistema" action={admin ? <button className="btn" onClick={addScreen}>+ anadir pantalla</button> : undefined}>
        <div className="systems-screen-grid">
          {screens.map((screen) => (
            <article className="systems-screen-card" key={screen.id}>
              <ImagePlaceholder
                admin={admin}
                image={screen.image}
                label="[captura de pantalla]"
                onImageChange={(image) => setScreens((current) => current.map((item) => item.id === screen.id ? { ...item, image } : item))}
              />
              <div>
                {admin && <button className="mini-delete" onClick={() => setScreens((current) => current.filter((item) => item.id !== screen.id))}>x</button>}
                <h3>
                  <EditableText
                    admin={admin}
                    value={screen.title}
                    onChange={(title) => setScreens((current) => current.map((item) => item.id === screen.id ? { ...item, title } : item))}
                  />
                </h3>
                <p>
                  <EditableText
                    admin={admin}
                    value={screen.description}
                    onChange={(description) => setScreens((current) => current.map((item) => item.id === screen.id ? { ...item, description } : item))}
                    multiline
                  />
                </p>
              </div>
            </article>
          ))}
        </div>
      </Section>

      <Section title="Pasos para probar" action={admin ? <button className="btn" onClick={addStep}>+ anadir paso</button> : undefined}>
        <div className="systems-step-list">
          {steps.map((step, index) => (
            <article className="systems-step" key={step.id}>
              <strong>{String(index + 1).padStart(2, '0')}</strong>
              <p>
                <EditableText
                  admin={admin}
                  value={step.text}
                  onChange={(text) => setSteps((current) => current.map((item) => item.id === step.id ? { ...item, text } : item))}
                  multiline
                />
              </p>
              {admin && <button className="btn danger" onClick={() => setSteps((current) => current.filter((item) => item.id !== step.id))}>x</button>}
            </article>
          ))}
        </div>
      </Section>
    </div>
  );
}

function CreativeProjectDetail({ admin, group, work }: { admin: boolean; group: ProjectGroupItem | null; work: GroupWorkItem | null }) {
  const scope = `${group?.id ?? 'creative'}-${work?.id ?? 'main'}`;
  const [title, setTitle] = usePersistentState(`portfolio.creative.${scope}.title`, work?.name ?? 'Identidad visual y piezas graficas');
  const [concept, setConcept] = usePersistentState(
    `portfolio.creative.${scope}.concept`,
    'Proyecto visual enfocado en construir una identidad clara, memorable y adaptable para piezas digitales e impresas.',
  );
  const [cover, setCover] = usePersistentState(`portfolio.creative.${scope}.cover`, '');
  const [client, setClient] = usePersistentState(`portfolio.creative.${scope}.client`, 'Marca / cliente');
  const [style, setStyle] = usePersistentState(`portfolio.creative.${scope}.style`, 'Minimalista, moderno y expresivo');
  const [pieces, setPieces] = usePersistentState<CreativePiece[]>(`portfolio.creative.${scope}.pieces`, [
    { id: 1, title: 'Logo principal', description: 'Version principal para uso digital e impreso.', image: '' },
    { id: 2, title: 'Piezas para redes', description: 'Composiciones adaptadas a publicaciones y formatos sociales.', image: '' },
    { id: 3, title: 'Aplicaciones de marca', description: 'Mockups para mostrar el sistema visual en contexto.', image: '' },
  ]);
  const [palette, setPalette] = usePersistentState<CreativeColor[]>(`portfolio.creative.${scope}.palette`, [
    { id: 1, name: 'Principal', value: '#45f0c1' },
    { id: 2, name: 'Contraste', value: '#f7a8ff' },
    { id: 3, name: 'Fondo', value: '#08060f' },
  ]);
  const [deliverables, setDeliverables] = usePersistentState<string[]>(`portfolio.creative.${scope}.deliverables`, [
    'Logotipo editable',
    'Manual breve de uso',
    'Piezas exportadas para redes',
  ]);
  const [process, setProcess] = usePersistentState<CreativeProcessStep[]>(`portfolio.creative.${scope}.process`, [
    { id: 1, title: 'Exploracion', description: 'Busqueda de referencias, tono visual y direccion grafica.' },
    { id: 2, title: 'Composicion', description: 'Creacion de propuestas, variantes y sistema de color.' },
    { id: 3, title: 'Entrega', description: 'Exportacion final, formatos y guia de uso.' },
  ]);

  const addPiece = () => setPieces((current) => [...current, { id: Date.now(), title: 'Nueva pieza', description: 'Describe esta pieza visual.', image: '' }]);
  const addColor = () => setPalette((current) => [...current, { id: Date.now(), name: 'Nuevo color', value: '#45f0c1' }]);
  const addDeliverable = () => setDeliverables((current) => [...current, 'Nuevo entregable']);
  const addProcessStep = () => setProcess((current) => [...current, { id: Date.now(), title: 'Nueva etapa', description: 'Describe esta parte del proceso.' }]);

  return (
    <div className="group-view group-view-creative">
      <p className="breadcrumb">Proyectos / {group?.name ?? 'Diseno grafico'}</p>
      <h1 className="title-line creative-title">
        <EditableText admin={admin} value={title} onChange={setTitle} />
        <EditMark show={admin} />
      </h1>

      <section className="creative-hero-panel">
        <div className="creative-cover">
          <ImagePlaceholder
            admin={admin}
            image={cover}
            label="[imagen principal del proyecto]"
            onImageChange={setCover}
          />
        </div>
        <div className="creative-brief">
          <span>Concepto visual</span>
          <p>
            <EditableText admin={admin} value={concept} onChange={setConcept} multiline />
            <EditMark show={admin} />
          </p>
          <div className="creative-meta-grid">
            <article>
              <small>CLIENTE</small>
              <strong><EditableText admin={admin} value={client} onChange={setClient} /></strong>
            </article>
            <article>
              <small>ESTILO</small>
              <strong><EditableText admin={admin} value={style} onChange={setStyle} /></strong>
            </article>
          </div>
        </div>
      </section>

      <Section title="Piezas visuales" action={admin ? <button className="btn" onClick={addPiece}>+ anadir pieza</button> : undefined}>
        <div className="creative-piece-grid">
          {pieces.map((piece) => (
            <article className="creative-piece-card" key={piece.id}>
              {admin && <button className="mini-delete" onClick={() => setPieces((current) => current.filter((item) => item.id !== piece.id))}>x</button>}
              <ImagePlaceholder
                admin={admin}
                image={piece.image}
                label="[pieza grafica]"
                onImageChange={(image) => setPieces((current) => current.map((item) => item.id === piece.id ? { ...item, image } : item))}
              />
              <h3>
                <EditableText
                  admin={admin}
                  value={piece.title}
                  onChange={(title) => setPieces((current) => current.map((item) => item.id === piece.id ? { ...item, title } : item))}
                />
              </h3>
              <p>
                <EditableText
                  admin={admin}
                  value={piece.description}
                  onChange={(description) => setPieces((current) => current.map((item) => item.id === piece.id ? { ...item, description } : item))}
                  multiline
                />
              </p>
            </article>
          ))}
        </div>
      </Section>

      <Section title="Paleta e identidad" action={admin ? <button className="btn" onClick={addColor}>+ anadir color</button> : undefined}>
        <div className="creative-palette-grid">
          {palette.map((color) => (
            <article className="creative-color-card" key={color.id}>
              {admin && <button className="mini-delete" onClick={() => setPalette((current) => current.filter((item) => item.id !== color.id))}>x</button>}
              <span style={{ background: color.value }} />
              <div>
                <strong>
                  <EditableText
                    admin={admin}
                    value={color.name}
                    onChange={(name) => setPalette((current) => current.map((item) => item.id === color.id ? { ...item, name } : item))}
                  />
                </strong>
                {admin ? (
                  <input
                    className="inline-field"
                    type="color"
                    value={color.value}
                    onChange={(event) => setPalette((current) => current.map((item) => item.id === color.id ? { ...item, value: event.target.value } : item))}
                  />
                ) : (
                  <small>{color.value}</small>
                )}
              </div>
            </article>
          ))}
        </div>
      </Section>

      <Section title="Proceso creativo" action={admin ? <button className="btn" onClick={addProcessStep}>+ anadir etapa</button> : undefined}>
        <div className="creative-process-list">
          {process.map((step, index) => (
            <article className="creative-process-step" key={step.id}>
              <strong>{String(index + 1).padStart(2, '0')}</strong>
              <div>
                <h3>
                  <EditableText
                    admin={admin}
                    value={step.title}
                    onChange={(title) => setProcess((current) => current.map((item) => item.id === step.id ? { ...item, title } : item))}
                  />
                </h3>
                <p>
                  <EditableText
                    admin={admin}
                    value={step.description}
                    onChange={(description) => setProcess((current) => current.map((item) => item.id === step.id ? { ...item, description } : item))}
                    multiline
                  />
                </p>
              </div>
              {admin && <button className="btn danger" onClick={() => setProcess((current) => current.filter((item) => item.id !== step.id))}>x</button>}
            </article>
          ))}
        </div>
      </Section>

      <Section title="Entregables" action={admin ? <button className="btn" onClick={addDeliverable}>+ anadir entregable</button> : undefined}>
        <div className="creative-deliverables">
          {deliverables.map((item, index) => (
            <article key={`${item}-${index}`}>
              <span>OK</span>
              <p>
                <EditableText
                  admin={admin}
                  value={item}
                  onChange={(value) => setDeliverables((current) => current.map((saved, savedIndex) => savedIndex === index ? value : saved))}
                />
              </p>
              {admin && <button className="btn danger" onClick={() => setDeliverables((current) => current.filter((_, savedIndex) => savedIndex !== index))}>x</button>}
            </article>
          ))}
        </div>
      </Section>
    </div>
  );
}

function ProjectDetail({ admin, group, work }: { admin: boolean; group: ProjectGroupItem | null; work: GroupWorkItem | null }) {
  const groupStyle = inferGroupStyle(group);
  if (groupStyle === 'systems') return <SystemsProjectDetail admin={admin} group={group} work={work} />;
  if (groupStyle === 'creative') return <CreativeProjectDetail admin={admin} group={group} work={work} />;

  return <DataProjectDetail admin={admin} />;
}

function DataProjectDetail({ admin }: { admin: boolean }) {
  const [projectTitle, setProjectTitle] = usePersistentState('portfolio.detail.title', 'Analisis de ventas y rentabilidad');
  const [managerOpen, setManagerOpen] = useState(true);
  const [viewerContext, setViewerContext] = useState<ViewerContext | null>(null);
  const [sectionLayout, setSectionLayout] = usePersistentState<DetailSectionConfig[]>('portfolio.detail.sectionLayout', getInitialSectionLayout());
  const [toolsList, setToolsList] = usePersistentState<DetailTool[]>('portfolio.detail.tools', [
    { id: 1, name: 'Python', description: 'Analisis y automatizacion', logo: 'PY', image: '' },
    { id: 2, name: 'Pandas', description: 'Limpieza de datos', logo: 'PD', image: '' },
    { id: 3, name: 'Matplotlib', description: 'Visualizacion', logo: 'MT', image: '' },
    { id: 4, name: 'Power BI', description: 'Dashboards', logo: 'BI', image: '' },
  ]);
  const [presentation, setPresentation] = usePersistentState(
    'portfolio.detail.presentation',
    'Analisis exhaustivo de las ventas de una empresa retail durante los ultimos 3 anos. Se identificaron patrones estacionales, productos mas rentables y segmentos de clientes con mayor valor.',
  );
  const [resources, setResources] = useHydratedPersistentState<StoredFile[]>('portfolio.detail.resources', ['BD POWER BI PRACTICAS.xlsx']);
  const [methodTitle, setMethodTitle] = usePersistentState('portfolio.detail.methodTitle', 'Metodologia KDD');
  const [processes, setProcesses] = useHydratedPersistentState<ProcessItem[]>('portfolio.detail.processes', [
    { id: 1, title: 'Limpieza de datos', description: 'Eliminacion de duplicados y valores nulos', files: ['BD POWER BI PRACTICAS.xlsx'] },
    { id: 2, title: 'Transformacion', description: 'Normalizacion y creacion de variables derivadas', files: [] },
  ]);
  const [processPreviews, setProcessPreviews] = useHydratedPersistentState<ArchitecturePreview[]>('portfolio.detail.processPreviews', []);
  const [architectureModels, setArchitectureModels] = useHydratedPersistentState<ArchitectureModel[]>('portfolio.detail.architectureModels', getInitialArchitectureModels());
  const [previews, setPreviews] = useHydratedPersistentState<StoredArchitecturePreview[]>('portfolio.detail.previews', []);
  const [queries, setQueries] = useHydratedPersistentState<QueryItem[]>('portfolio.detail.queries', [
    {
      id: 1,
      question: 'Cual es el mes con mayor volumen de ventas?',
      answer: 'Diciembre concentra el 23% de las ventas anuales.',
      preview: null,
    },
  ]);

  const canShow = (section: DetailSectionConfig) => admin || section.visible;
  const updateSectionTitle = (id: number, title: string) => {
    setSectionLayout((current) => current.map((section) => section.id === id ? { ...section, title } : section));
  };
  const toggleSection = (id: number) => {
    setSectionLayout((current) => current.map((section) => section.id === id ? { ...section, visible: !section.visible } : section));
  };
  const moveSection = (id: number, direction: -1 | 1) => {
    setSectionLayout((current) => {
      const index = current.findIndex((section) => section.id === id);
      const nextIndex = index + direction;
      if (index < 0 || nextIndex < 0 || nextIndex >= current.length) return current;
      const next = [...current];
      [next[index], next[nextIndex]] = [next[nextIndex], next[index]];
      return next;
    });
  };
  const addSection = (type: DetailSectionKey) => {
    setSectionLayout((current) => [...current, { id: Date.now(), type, title: sectionTypeLabels[type], visible: true }]);
  };
  const removeSection = (id: number) => {
    setSectionLayout((current) => current.filter((section) => section.id !== id));
  };
  const addTool = () => {
    setToolsList((current) => [...current, { id: Date.now(), name: 'Nueva herramienta', description: 'Descripcion breve', logo: 'NH', image: '' }]);
  };
  const updateTool = (id: number, patch: Partial<DetailTool>) => {
    setToolsList((current) => current.map((item) => item.id === id ? { ...item, ...patch } : item));
  };
  const uploadDetailToolLogo = (id: number, file: File | undefined) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateTool(id, { image: String(reader.result) });
    reader.readAsDataURL(file);
  };
  const openFileViewer = (file: StoredFile, previewSourceKey?: string) => {
    setViewerContext({ file, previewSourceKey });
  };
  const viewerProcessPreviews = viewerContext?.previewSourceKey
    ? processPreviews.filter((preview) => preview.sourceKey === viewerContext.previewSourceKey)
    : [];
  const renderDetailSection = (section: DetailSectionConfig) => {
    if (!canShow(section)) return null;

    if (section.type === 'tools') {
      return (
        <Section key={section.id} title={section.title} action={admin ? <button className="btn" onClick={addTool}>+ anadir herramienta</button> : undefined}>
          <div className="detail-tools-grid">
            {toolsList.map((tool) => (
              <article className={admin ? 'detail-tool-card admin-card' : 'detail-tool-card'} key={tool.id}>
                {admin && (
                  <div className="tool-admin-actions">
                    <label className="btn upload-button tiny">
                      logo
                      <input type="file" accept="image/*" onChange={(event) => uploadDetailToolLogo(tool.id, event.target.files?.[0])} />
                    </label>
                    <button className="mini-delete" onClick={() => setToolsList((current) => current.filter((item) => item.id !== tool.id))}>x</button>
                  </div>
                )}
                <span className="detail-tool-icon">
                  {tool.image ? <img src={tool.image} alt={`Logo ${tool.name}`} /> : tool.logo}
                </span>
                <div className="detail-tool-copy">
                  <h3>
                    <EditableText admin={admin} value={tool.name} onChange={(name) => updateTool(tool.id, { name, logo: name.slice(0, 2).toUpperCase() || tool.logo })} />
                    <EditMark show={admin} />
                  </h3>
                  {(admin || tool.description) && (
                    <p>
                      <EditableText admin={admin} value={tool.description} onChange={(description) => updateTool(tool.id, { description })} />
                    </p>
                  )}
                </div>
              </article>
            ))}
          </div>
        </Section>
      );
    }

    if (section.type === 'presentation') {
      return (
        <Section key={section.id} title={section.title}>
          <div className="text-block editable-block">
            <EditableText admin={admin} value={presentation} onChange={setPresentation} multiline />
            <EditMark show={admin} />
          </div>
        </Section>
      );
    }

    if (section.type === 'resources') {
      return (
        <Section
          key={section.id}
          title={section.title}
          action={admin ? (
            <FileUploadButton
              label="+ anadir archivo o recurso"
              onUpload={(asset) => setResources((current) => [...current, asset])}
            />
          ) : undefined}
        >
          {resources.map((file) => (
            <FileRow
              key={fileKey(file)}
              admin={admin}
              file={file}
              onView={openFileViewer}
              onRename={(name) => setResources((current) => current.map((item) => item === file ? renameStoredFile(item, name) : item))}
              onDelete={() => setResources((current) => current.filter((item) => item !== file))}
            />
          ))}
        </Section>
      );
    }

    if (section.type === 'processes') {
      return (
        <ProcessSection
          key={section.id}
          admin={admin}
          title={section.title}
          methodTitle={methodTitle}
          setMethodTitle={setMethodTitle}
          processes={processes}
          setProcesses={setProcesses}
          processPreviews={processPreviews}
          setProcessPreviews={setProcessPreviews}
          onViewFile={openFileViewer}
        />
      );
    }

    if (section.type === 'architecture') {
      return (
        <ArchitectureSection
          key={section.id}
          admin={admin}
          title={section.title}
          models={architectureModels}
          setModels={setArchitectureModels}
          previews={previews}
          setPreviews={setPreviews}
          onViewFile={openFileViewer}
        />
      );
    }

    return (
      <QueriesSection
        key={section.id}
        admin={admin}
        title={section.title}
        queries={queries}
        setQueries={setQueries}
        onViewFile={openFileViewer}
      />
    );
  };
  return (
    <>
      <p className="breadcrumb">Proyectos / Analisis de Datos y Business Intelligence</p>
      <h1 className="title-line">
        <EditableText admin={admin} value={projectTitle} onChange={setProjectTitle} />
        <EditMark show={admin} />
      </h1>
      {admin && (
        <>
          <button className="manage" onClick={() => setManagerOpen((open) => !open)}>
            * GESTIONAR SECCIONES <span>{managerOpen ? '^' : 'v'}</span>
          </button>
          {managerOpen && (
            <ManageSections
              sections={sectionLayout}
              onAdd={addSection}
              onMove={moveSection}
              onRemove={removeSection}
              onToggle={toggleSection}
              onEdit={updateSectionTitle}
            />
          )}
        </>
      )}

      {sectionLayout.map(renderDetailSection)}
      {viewerContext && (
        <FileViewerModal
          admin={admin}
          file={viewerContext.file}
          fallbackPreviews={viewerProcessPreviews}
          onAddFallbackPreview={viewerContext.previewSourceKey ? (asset) => setProcessPreviews((current) => [
            ...current,
            { id: Date.now(), sourceKey: viewerContext.previewSourceKey ?? '', file: asset },
          ]) : undefined}
          onClose={() => setViewerContext(null)}
          onRemoveFallbackPreview={viewerContext.previewSourceKey ? (id) => setProcessPreviews((current) => current.filter((preview) => preview.id !== id)) : undefined}
        />
      )}
    </>
  );
}

function ManageSections({
  sections,
  onAdd,
  onMove,
  onRemove,
  onToggle,
  onEdit,
}: {
  sections: DetailSectionConfig[];
  onAdd: (type: DetailSectionKey) => void;
  onMove: (id: number, direction: -1 | 1) => void;
  onRemove: (id: number) => void;
  onToggle: (id: number) => void;
  onEdit: (id: number, title: string) => void;
}) {
  const [newSectionType, setNewSectionType] = useState<DetailSectionKey>('presentation');
  return (
    <div className="section-manager">
      <div className="manager-list">
        {sections.map((section, index) => (
          <div className="manager-item" key={section.id}>
            <div className="manager-order">
              <button className="btn" disabled={index === 0} onClick={() => onMove(section.id, -1)}>^</button>
              <button className="btn" disabled={index === sections.length - 1} onClick={() => onMove(section.id, 1)}>v</button>
            </div>
            <button className="manager-visible" onClick={() => onToggle(section.id)}>{section.visible ? 'o' : '-'}</button>
            <span className="manager-kind">{sectionTypeLabels[section.type]}</span>
            <EditableText admin value={section.title} onChange={(title) => onEdit(section.id, title)} />
            <button className="btn danger" onClick={() => onRemove(section.id)}>x</button>
          </div>
        ))}
      </div>
      <div className="manager-add-row">
        <select value={newSectionType} onChange={(event) => setNewSectionType(event.target.value as DetailSectionKey)}>
          {sectionTypes.map((type) => (
            <option key={type} value={type}>{sectionTypeLabels[type]}</option>
          ))}
        </select>
        <button className="btn" onClick={() => onAdd(newSectionType)}>+ anadir seccion</button>
      </div>
      <p>o = visible para visitantes - - = oculto para visitantes - usa ^ / v para ordenar - el titulo es editable</p>
    </div>
  );
}

function FileUploadButton({ label, onUpload, accept }: { label: string; onUpload: (asset: FileAsset) => void; accept?: string }) {
  return (
    <label className="btn upload-button file-upload-button">
      {label}
      <input type="file" accept={accept} onChange={(event) => {
        const file = event.target.files?.[0];
        if (file) readFileAsset(file, onUpload);
        event.currentTarget.value = '';
      }} />
    </label>
  );
}

function FileRow({
  admin,
  file = 'BD POWER BI PRACTICAS.xlsx',
  onView,
  onRename,
  onDelete,
}: {
  admin: boolean;
  file?: StoredFile;
  onView?: (file: StoredFile) => void;
  onRename?: (name: string) => void;
  onDelete?: () => void;
}) {
  const name = fileName(file);

  return (
    <div className="file-row">
      <div className="file-meta">
        <EditableText admin={admin && Boolean(onRename)} value={name} onChange={(value) => onRename?.(value)} />
      </div>
      <div>
        <button className="btn" onClick={() => onView?.(file)}>ver</button>
        {admin && <button className="btn danger" onClick={onDelete}>x</button>}
      </div>
    </div>
  );
}

function PreviewChoice({
  source,
  onUseSource,
  onUploadPreview,
}: {
  source: StoredFile;
  onUseSource: () => void;
  onUploadPreview: (asset: FileAsset) => void;
}) {
  const [open, setOpen] = useState(false);

  return (
    <div className="preview-choice">
      <button className="btn" onClick={() => setOpen((current) => !current)}>+ vista previa</button>
      {open && (
        <div className="preview-choice-menu">
          <button className="btn" onClick={() => {
            onUseSource();
            setOpen(false);
          }}>
            usar este archivo
          </button>
          <FileUploadButton
            label="+ subir captura/imagen"
            onUpload={(asset) => {
              onUploadPreview(asset);
              setOpen(false);
            }}
          />
        </div>
      )}
      <small>{fileName(source)}</small>
    </div>
  );
}

function FilePreview({
  file,
  admin,
  onUpload,
  onView,
}: {
  file: StoredFile | null;
  admin: boolean;
  onUpload: (asset: FileAsset) => void;
  onView?: (file: StoredFile) => void;
}) {
  return (
    <div className="file-preview">
      {file && isImageFile(file) && typeof file !== 'string' ? (
        <button className="preview-image-button" onClick={() => onView?.(file)}>
          <img src={file.url} alt={file.name} />
        </button>
      ) : file ? (
        <button className="preview-file" onClick={() => onView?.(file)}>
          <span>{fileName(file)}</span>
          <small>{typeof file === 'string' ? 'archivo pendiente' : file.type || 'archivo'}</small>
        </button>
      ) : (
        <p>[sin vista previa]</p>
      )}
      {admin && (
        <div className="preview-actions">
          <FileUploadButton label="+ subir vista previa" onUpload={onUpload} />
        </div>
      )}
    </div>
  );
}

function PreviewTile({
  file,
  admin,
  onView,
  onRemove,
}: {
  file: StoredFile;
  admin: boolean;
  onView: (file: StoredFile) => void;
  onRemove?: () => void;
}) {
  const [rows, setRows] = useState<string[][]>([]);

  useEffect(() => {
    setRows([]);
    if (!isExcelFile(file) || typeof file === 'string') return;

    fetch(file.url)
      .then((response) => response.arrayBuffer())
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const table = XLSX.utils.sheet_to_json<Array<string | number | boolean | null>>(sheet, {
          header: 1,
          defval: '',
        });
        setRows(table.slice(0, 6).map((row) => row.slice(0, 5).map((cell) => String(cell ?? ''))));
      })
      .catch(() => setRows([]));
  }, [file]);

  return (
    <div className="preview-tile">
      <button className="preview-tile-view" onClick={() => onView(file)}>
        {isImageFile(file) && typeof file !== 'string' ? (
          <img src={file.url} alt={file.name} />
        ) : isPdfFile(file) ? (
          <iframe src={typeof file === 'string' ? undefined : file.url} title={fileName(file)} />
        ) : rows.length > 0 ? (
          <table className="preview-mini-table">
            <tbody>
              {rows.map((row, rowIndex) => (
                <tr key={`preview-row-${rowIndex}`}>
                  {row.map((cell, cellIndex) => (
                    <td key={`preview-cell-${rowIndex}-${cellIndex}`}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <span>
            {isExcelFile(file) ? 'EXCEL' : isPowerBiProject(file) ? 'PBIX' : 'ARCHIVO'}
            <small>{fileName(file)}</small>
          </span>
        )}
      </button>
      {admin && <button className="preview-remove btn danger" onClick={onRemove}>x</button>}
    </div>
  );
}

function PreviewCarousel({
  previews,
  admin,
  onView,
  onRemove,
}: {
  previews: ArchitecturePreview[];
  admin: boolean;
  onView: (file: StoredFile) => void;
  onRemove: (id: number) => void;
}) {
  const [activeIndex, setActiveIndex] = useState(0);

  useEffect(() => {
    if (activeIndex > previews.length - 1) setActiveIndex(Math.max(0, previews.length - 1));
  }, [activeIndex, previews.length]);

  const current = previews[activeIndex];

  if (!current) return null;

  return (
    <div className="preview-carousel-shell">
      <button className="preview-nav left" onClick={() => setActiveIndex((index) => Math.max(0, index - 1))}>‹</button>
      <PreviewTile
        admin={admin}
        file={current.file}
        onView={onView}
        onRemove={() => onRemove(current.id)}
      />
      <button className="preview-nav right" onClick={() => setActiveIndex((index) => Math.min(previews.length - 1, index + 1))}>›</button>
      {previews.length > 1 && <span className="preview-count">{activeIndex + 1} / {previews.length}</span>}
    </div>
  );
}

function PreviewPicker({
  files,
  previews,
  onAdd,
}: {
  files: StoredFile[];
  previews: StoredFile[];
  onAdd: (file: StoredFile) => void;
}) {
  const availableFiles = files.filter((file) => !previews.some((preview) => sameStoredFile(preview, file)));
  const [selectedKey, setSelectedKey] = useState('');

  useEffect(() => {
    setSelectedKey('');
  }, [files.length, previews.length]);

  if (availableFiles.length === 0) {
    return <p className="empty">Sube archivos en el bloque de la izquierda para poder agregarlos como vista previa.</p>;
  }

  return (
    <div className="preview-picker">
      <select value={selectedKey} onChange={(event) => setSelectedKey(event.target.value)}>
        <option value="">Seleccionar archivo</option>
        {availableFiles.map((file) => (
          <option key={fileKey(file)} value={fileKey(file)}>
            {fileName(file)}
          </option>
        ))}
      </select>
      <button
        className="btn"
        onClick={() => {
          const selected = availableFiles.find((file) => fileKey(file) === selectedKey);
          if (selected) onAdd(selected);
        }}
      >
        + anadir vista previa
      </button>
    </div>
  );
}

function FileViewerModal({
  admin = false,
  file,
  fallbackPreviews = [],
  onAddFallbackPreview,
  onClose,
  onRemoveFallbackPreview,
}: {
  admin?: boolean;
  file: StoredFile;
  fallbackPreviews?: ArchitecturePreview[];
  onAddFallbackPreview?: (asset: FileAsset) => void;
  onClose: () => void;
  onRemoveFallbackPreview?: (id: number) => void;
}) {
  const [textContent, setTextContent] = useState('');
  const [workbookSheets, setWorkbookSheets] = useState<Array<{ name: string; rows: string[][] }>>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [error, setError] = useState('');

  useEffect(() => {
    setTextContent('');
    setWorkbookSheets([]);
    setActiveSheet(0);
    setError('');

    if (typeof file === 'string') return;

    if (isExcelFile(file)) {
      fetch(file.url)
        .then((response) => response.arrayBuffer())
        .then((buffer) => {
          const workbook = XLSX.read(buffer, { type: 'array' });
          const sheets = workbook.SheetNames.map((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json<Array<string | number | boolean | null>>(sheet, {
              header: 1,
              defval: '',
            });
            return {
              name: sheetName,
              rows: rows.map((row) => row.map((cell) => String(cell ?? ''))),
            };
          });
          setWorkbookSheets(sheets);
        })
        .catch(() => setError('No se pudo leer este archivo Excel.'));
      return;
    }

    if (isTextFile(file)) {
      fetch(file.url)
        .then((response) => response.text())
        .then(setTextContent)
        .catch(() => setError('No se pudo leer este archivo de texto.'));
    }
  }, [file]);

  const name = fileName(file);
  const canRenderFrame = typeof file !== 'string' && (isPdfFile(file) || file.type === 'text/html');
  const isWaitingForParsedContent = typeof file !== 'string' && (isExcelFile(file) || isTextFile(file));
  const canUseInternalViewer = typeof file !== 'string' && (isImageFile(file) || isExcelFile(file) || isTextFile(file) || canRenderFrame);
  const canUseFallbackPreviews = !canUseInternalViewer || Boolean(error);
  const fallbackPanel = canUseFallbackPreviews && (fallbackPreviews.length > 0 || (admin && onAddFallbackPreview)) ? (
    <div className="viewer-fallback-panel">
      <span>VISTAS ANADIDAS</span>
      {fallbackPreviews.length > 0 ? (
        <PreviewCarousel
          admin={admin}
          previews={fallbackPreviews}
          onView={() => undefined}
          onRemove={(id) => onRemoveFallbackPreview?.(id)}
        />
      ) : (
        <div className="viewer-empty compact"><p>[sin vista alternativa]</p></div>
      )}
      {admin && onAddFallbackPreview && (
        <div className="preview-actions">
          <FileUploadButton label="+ anadir captura o imagen" accept="image/*" onUpload={onAddFallbackPreview} />
        </div>
      )}
    </div>
  ) : null;

  return (
    <div className="viewer-backdrop" role="dialog" aria-modal="true">
      <section className="viewer-modal">
        <header>
          <div>
            <span>VISTA DE ARCHIVO</span>
            <h2>{name}</h2>
          </div>
          <button className="btn danger" onClick={onClose}>x</button>
        </header>
        <div className="viewer-body">
          {fallbackPreviews.length > 0 && canUseFallbackPreviews ? (
            fallbackPanel
          ) : typeof file === 'string' ? (
            <div className="viewer-empty">
              <p>Este archivo solo tiene nombre guardado. Vuelve a subirlo con el boton de anadir archivo para poder ver su contenido aqui.</p>
              {fallbackPanel}
            </div>
          ) : isImageFile(file) ? (
            <img className="viewer-image" src={file.url} alt={file.name} />
          ) : workbookSheets.length > 0 ? (
            <div className="viewer-workbook">
              <div className="viewer-sheet-tabs">
                {workbookSheets.map((sheet, index) => (
                  <button
                    className={index === activeSheet ? 'active' : ''}
                    key={sheet.name}
                    onClick={() => setActiveSheet(index)}
                  >
                    {sheet.name}
                  </button>
                ))}
              </div>
              <div className="viewer-table-wrap">
                <table className="viewer-table">
                  <tbody>
                    {workbookSheets[activeSheet].rows.slice(0, 120).map((row, rowIndex) => (
                      <tr key={`row-${rowIndex}`}>
                        {row.slice(0, 32).map((cell, cellIndex) => (
                          <td key={`cell-${rowIndex}-${cellIndex}`}>{cell}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : textContent ? (
            <pre className="viewer-text">{textContent}</pre>
          ) : canRenderFrame ? (
            <iframe className="viewer-frame" src={file.url} title={file.name} />
          ) : error ? (
            <div className="viewer-empty"><p>{error}</p>{fallbackPanel}</div>
          ) : isWaitingForParsedContent ? (
            <div className="viewer-empty"><p>Cargando vista previa...</p></div>
          ) : isPowerBiProject(file) ? (
            <div className="viewer-empty">
              <p>El archivo .pbix queda guardado como recurso, pero por ahora no se puede renderizar directamente en el navegador. Luego vemos la vista de Power BI.</p>
              {fallbackPanel}
            </div>
          ) : (
            <div className="viewer-empty"><p>No hay visor interno disponible para este tipo de archivo. Puedes subir una captura o imagen para verlo aqui.</p>{fallbackPanel}</div>
          )}
        </div>
      </section>
    </div>
  );
}

function ProcessSection({
  admin,
  title,
  methodTitle,
  setMethodTitle,
  processes,
  setProcesses,
  processPreviews,
  setProcessPreviews,
  onViewFile,
}: {
  admin: boolean;
  title: string;
  methodTitle: string;
  setMethodTitle: React.Dispatch<React.SetStateAction<string>>;
  processes: ProcessItem[];
  setProcesses: React.Dispatch<React.SetStateAction<ProcessItem[]>>;
  processPreviews: ArchitecturePreview[];
  setProcessPreviews: React.Dispatch<React.SetStateAction<ArchitecturePreview[]>>;
  onViewFile: (file: StoredFile, previewSourceKey?: string) => void;
}) {
  const addProcess = () => {
    setProcesses((current) => [...current, { id: Date.now(), title: 'Nuevo proceso', description: 'Descripcion del proceso', files: [] }]);
  };
  const addProcessFile = (id: number, asset: FileAsset) => {
    setProcesses((current) => current.map((process) => process.id === id ? { ...process, files: [...process.files, asset] } : process));
  };

  return (
    <Section title={title} action={admin ? <button className="btn" onClick={addProcess}>+ anadir grupo de procesos</button> : undefined}>
      <div className="process-group">
        <p>
          METODO:{' '}
          <u><EditableText admin={admin} value={methodTitle} onChange={setMethodTitle} /></u>{' '}
          <EditMark show={admin} />
        </p>
        <div className="process-grid">
          {processes.map((process) => (
            <article className="process-card" key={process.id}>
              {admin && <button className="delete" onClick={() => {
                setProcesses((current) => current.filter((item) => item.id !== process.id));
                setProcessPreviews((current) => current.filter((preview) => !process.files.some((file) => preview.sourceKey === processSourceKey(process.id, file))));
              }}>x</button>}
              <h3>
                <EditableText
                  admin={admin}
                  value={process.title}
                  onChange={(title) => setProcesses((current) => current.map((item) => item.id === process.id ? { ...item, title } : item))}
                />{' '}
                <EditMark show={admin} />
              </h3>
              <p>
                <EditableText
                  admin={admin}
                  value={process.description}
                  onChange={(description) => setProcesses((current) => current.map((item) => item.id === process.id ? { ...item, description } : item))}
                />{' '}
                <EditMark show={admin} />
              </p>
              {process.files.map((file) => (
                <FileRow
                  admin={admin}
                  key={fileKey(file)}
                  file={file}
                  onView={(savedFile) => onViewFile(savedFile, processSourceKey(process.id, savedFile))}
                  onRename={(name) => setProcesses((current) => current.map((item) => item.id === process.id ? { ...item, files: item.files.map((saved) => saved === file ? renameStoredFile(saved, name) : saved) } : item))}
                  onDelete={() => {
                    setProcesses((current) => current.map((item) => item.id === process.id ? { ...item, files: item.files.filter((saved) => saved !== file) } : item));
                    setProcessPreviews((current) => current.filter((preview) => preview.sourceKey !== processSourceKey(process.id, file)));
                  }}
                />
              ))}
              {admin && <FileUploadButton label="+ anadir archivo" onUpload={(asset) => addProcessFile(process.id, asset)} />}
            </article>
          ))}
          {admin && <button className="add-process" onClick={addProcess}>+ anadir proceso</button>}
        </div>
      </div>
    </Section>
  );
}

function ArchitectureSection({
  admin,
  title,
  models,
  setModels,
  previews,
  setPreviews,
  onViewFile,
}: {
  admin: boolean;
  title: string;
  models: ArchitectureModel[];
  setModels: React.Dispatch<React.SetStateAction<ArchitectureModel[]>>;
  previews: StoredArchitecturePreview[];
  setPreviews: React.Dispatch<React.SetStateAction<StoredArchitecturePreview[]>>;
  onViewFile: (file: StoredFile) => void;
}) {
  const allArchitectureFiles = models.flatMap((model) => model.segments.flatMap((segment) => segment.files));
  const normalizedPreviews = previews.map((preview, index) => {
    if (isArchitecturePreview(preview)) return preview;
    const source = allArchitectureFiles.find((file) => sameStoredFile(file, preview)) ?? allArchitectureFiles[0];
    return {
      id: Date.now() + index,
      sourceKey: source ? fileKey(source) : '',
      file: preview,
    };
  }).filter((preview) => preview.sourceKey);
  const addModel = () => {
    setModels((current) => [
      ...current,
      {
        id: Date.now(),
        title: 'Nuevo modelado',
        segments: [{ id: Date.now() + 1, title: 'Nuevo segmento', description: 'Descripcion del segmento', files: [] }],
      },
    ]);
  };
  const updateModel = (modelId: number, patch: Partial<ArchitectureModel>) => {
    setModels((current) => current.map((model) => model.id === modelId ? { ...model, ...patch } : model));
  };
  const deleteModel = (model: ArchitectureModel) => {
    setModels((current) => current.filter((item) => item.id !== model.id));
    setPreviews((current) => current.filter((preview) => {
      const sourceKey = isArchitecturePreview(preview) ? preview.sourceKey : fileKey(preview);
      return !model.segments.some((segment) => segment.files.some((file) => matchesArchitectureSourceKey(sourceKey, model.id, segment.id, file)));
    }));
  };
  const addSegment = (modelId: number) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: [...model.segments, { id: Date.now(), title: 'Nuevo segmento', description: 'Descripcion del segmento', files: [] }],
    } : model));
  };
  const updateSegment = (modelId: number, segmentId: number, patch: Partial<ProcessItem>) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: model.segments.map((segment) => segment.id === segmentId ? { ...segment, ...patch } : segment),
    } : model));
  };
  const deleteSegment = (modelId: number, segment: ProcessItem) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: model.segments.filter((item) => item.id !== segment.id),
    } : model));
    setPreviews((current) => current.filter((preview) => {
      const sourceKey = isArchitecturePreview(preview) ? preview.sourceKey : fileKey(preview);
      return !segment.files.some((file) => matchesArchitectureSourceKey(sourceKey, modelId, segment.id, file));
    }));
  };
  const addArchitectureFile = (modelId: number, segmentId: number, asset: FileAsset) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: model.segments.map((segment) => segment.id === segmentId ? {
        ...segment,
        files: [...segment.files, asset],
      } : segment),
    } : model));
  };
  const renameArchitectureFile = (modelId: number, segmentId: number, file: StoredFile, name: string) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: model.segments.map((segment) => segment.id === segmentId ? {
        ...segment,
        files: segment.files.map((saved) => sameStoredFile(saved, file) ? renameStoredFile(saved, name) : saved),
      } : segment),
    } : model));
  };
  const deleteArchitectureFile = (modelId: number, segmentId: number, file: StoredFile) => {
    setModels((current) => current.map((model) => model.id === modelId ? {
      ...model,
      segments: model.segments.map((segment) => segment.id === segmentId ? {
        ...segment,
        files: segment.files.filter((saved) => !sameStoredFile(saved, file)),
      } : segment),
    } : model));
    setPreviews((current) => current.filter((preview) => {
      const sourceKey = isArchitecturePreview(preview) ? preview.sourceKey : fileKey(preview);
      return !matchesArchitectureSourceKey(sourceKey, modelId, segmentId, file);
    }));
  };
  const previewsFor = (modelId: number, segmentId: number, source: StoredFile) => (
    normalizedPreviews.filter((preview) => matchesArchitectureSourceKey(preview.sourceKey, modelId, segmentId, source))
  );
  const addPreviewFor = (modelId: number, segmentId: number, source: StoredFile, file: StoredFile) => {
    setPreviews((current) => [
      ...current,
      { id: Date.now(), sourceKey: architectureSourceKey(modelId, segmentId, source), file },
    ]);
  };

  return (
    <Section title={title} action={admin ? <button className="btn" onClick={addModel}>+ anadir modelado</button> : undefined}>
      <div className="architecture-model-list">
        {models.map((model) => (
          <article className="architecture-model" key={model.id}>
            {admin && <button className="delete" onClick={() => deleteModel(model)}>x modelado</button>}
            <p>
              MODELADO:{' '}
              <u><EditableText admin={admin} value={model.title} onChange={(modelTitle) => updateModel(model.id, { title: modelTitle })} /></u>{' '}
              <EditMark show={admin} />
            </p>
            <div className="architecture-segment-list">
              {model.segments.map((segment) => (
                <article className="architecture-segment" key={segment.id}>
                  {admin && <button className="delete" onClick={() => deleteSegment(model.id, segment)}>x segmento</button>}
                  <h3>
                    <EditableText admin={admin} value={segment.title} onChange={(segmentTitle) => updateSegment(model.id, segment.id, { title: segmentTitle })} />{' '}
                    <EditMark show={admin} />
                  </h3>
                  <p>
                    <EditableText admin={admin} value={segment.description} onChange={(description) => updateSegment(model.id, segment.id, { description })} />{' '}
                    <EditMark show={admin} />
                  </p>
                  <div className="architecture-files">
                    {segment.files.map((file) => {
                      const filePreviews = previewsFor(model.id, segment.id, file);
                      return (
                        <div className="architecture-file-row" key={fileKey(file)}>
                          <div className="architecture-file-header">
                            <FileRow
                              admin={admin}
                              file={file}
                              onView={onViewFile}
                              onRename={(name) => renameArchitectureFile(model.id, segment.id, file, name)}
                              onDelete={() => deleteArchitectureFile(model.id, segment.id, file)}
                            />
                            {admin && (
                              <PreviewChoice
                                source={file}
                                onUseSource={() => addPreviewFor(model.id, segment.id, file, file)}
                                onUploadPreview={(asset) => addPreviewFor(model.id, segment.id, file, asset)}
                              />
                            )}
                          </div>
                          {filePreviews.length > 0 && (
                            <div className="architecture-preview-stage">
                              <PreviewCarousel
                                admin={admin}
                                previews={filePreviews}
                                onView={onViewFile}
                                onRemove={(id) => setPreviews((current) => current.filter((item) => !(isArchitecturePreview(item) && item.id === id)))}
                              />
                            </div>
                          )}
                        </div>
                      );
                    })}
                    {segment.files.length === 0 && <div className="architecture-empty"><p>[sin archivos]</p></div>}
                  </div>
                  {admin && (
                    <div className="preview-actions left-actions">
                      <FileUploadButton label="+ anadir archivo" onUpload={(asset) => addArchitectureFile(model.id, segment.id, asset)} />
                    </div>
                  )}
                </article>
              ))}
              {model.segments.length === 0 && <div className="architecture-empty"><p>[sin segmentos]</p></div>}
            </div>
            {admin && <button className="btn" onClick={() => addSegment(model.id)}>+ anadir segmento</button>}
          </article>
        ))}
        {models.length === 0 && <div className="architecture-empty"><p>[sin modelados]</p></div>}
      </div>
    </Section>
  );
}

function QueriesSection({
  admin,
  title,
  queries,
  setQueries,
  onViewFile,
}: {
  admin: boolean;
  title: string;
  queries: QueryItem[];
  setQueries: React.Dispatch<React.SetStateAction<QueryItem[]>>;
  onViewFile: (file: StoredFile) => void;
}) {
  const [openQueries, setOpenQueries] = useState<Record<number, boolean>>({});
  const addQuery = () => {
    setQueries((current) => [...current, { id: Date.now(), question: 'Nueva pregunta', answer: 'Nueva respuesta', previews: [] }]);
  };
  const isQueryOpen = (queryId: number) => openQueries[queryId] ?? true;
  const toggleQuery = (queryId: number) => {
    setOpenQueries((current) => ({ ...current, [queryId]: !(current[queryId] ?? true) }));
  };
  const getQueryPreviews = (query: QueryItem) => {
    const savedPreviews = query.previews ?? [];
    const legacyPreview = query.preview ? [{ id: query.id, sourceKey: `query-${query.id}`, file: query.preview }] : [];

    return [...legacyPreview, ...savedPreviews.map((preview, index) => (
      isArchitecturePreview(preview)
        ? preview
        : { id: query.id + index + 1, sourceKey: `query-${query.id}`, file: preview }
    ))];
  };
  const addQueryPreview = (queryId: number, preview: FileAsset) => {
    setQueries((current) => current.map((item) => item.id === queryId ? {
      ...item,
      preview: null,
      previews: [...(item.previews ?? []), { id: Date.now(), sourceKey: `query-${queryId}`, file: preview }],
    } : item));
  };
  const removeQueryPreview = (queryId: number, previewId: number) => {
    setQueries((current) => current.map((item) => item.id === queryId ? {
      ...item,
      preview: item.preview && previewId === item.id ? null : item.preview,
      previews: (item.previews ?? []).filter((preview, index) => {
        const id = isArchitecturePreview(preview) ? preview.id : item.id + index + 1;
        return id !== previewId;
      }),
    } : item));
  };

  return (
    <Section title={title} action={admin ? <button className="btn" onClick={addQuery}>+ anadir consulta</button> : undefined}>
      {queries.map((query, index) => (
        <div className="query-box" key={query.id}>
          <div className="query-head">
            <span>[pregunta {index + 1}]</span>
            <p>
              <EditableText
                admin={admin}
                value={query.question}
                onChange={(question) => setQueries((current) => current.map((item) => item.id === query.id ? { ...item, question } : item))}
              />
            </p>
            {admin && <button className="btn danger" onClick={() => setQueries((current) => current.filter((item) => item.id !== query.id))}>x</button>}
            <button className="plain query-toggle" onClick={() => toggleQuery(query.id)}>{isQueryOpen(query.id) ? '^' : 'v'}</button>
          </div>
          {isQueryOpen(query.id) && (
            <div className="query-detail">
              <div className="qa-column">
                <article>
                  <span>PREGUNTA</span>
                  <p>
                    <EditableText
                      admin={admin}
                      value={query.question}
                      onChange={(question) => setQueries((current) => current.map((item) => item.id === query.id ? { ...item, question } : item))}
                      multiline
                    />{' '}
                    <EditMark show={admin} />
                  </p>
                </article>
                <article className="answer">
                  <span>RESPUESTA</span>
                  <p>
                    <EditableText
                      admin={admin}
                      value={query.answer}
                      onChange={(answer) => setQueries((current) => current.map((item) => item.id === query.id ? { ...item, answer } : item))}
                      multiline
                    />{' '}
                    <EditMark show={admin} />
                  </p>
                </article>
              </div>
              <div className="query-preview">
                <span>VISTA PREVIA</span>
                {getQueryPreviews(query).length > 0 ? (
                  <div className="query-preview-carousel">
                    <PreviewCarousel
                      admin={admin}
                      previews={getQueryPreviews(query)}
                      onView={onViewFile}
                      onRemove={(id) => removeQueryPreview(query.id, id)}
                    />
                  </div>
                ) : (
                  <div className="preview-empty">
                    <p>[sin vista previa]</p>
                  </div>
                )}
                {admin && (
                  <div className="preview-actions">
                    <FileUploadButton label="+ subir vista previa" onUpload={(preview) => addQueryPreview(query.id, preview)} />
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      ))}
    </Section>
  );
}

function Dashboard({
  role,
  onLogout,
  theme,
  setTheme,
}: {
  role: Role;
  onLogout: () => void;
  theme: PortfolioTheme;
  setTheme: React.Dispatch<React.SetStateAction<PortfolioTheme>>;
}) {
  const [view, setView] = useState<View>('home');
  const [selectedGroup, setSelectedGroup] = useState<ProjectGroupItem | null>(null);
  const [selectedWork, setSelectedWork] = useState<GroupWorkItem | null>(null);
  const [displayMode, setDisplayMode] = usePersistentState<'day' | 'night'>('portfolio.displayMode', 'night');
  const admin = role === 'admin';

  const navigate = (target: SectionId) => {
    setView(target);
    window.setTimeout(() => {
      document.getElementById(target)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 0);
  };

  return (
    <main className={`app-shell app-theme-data theme-${displayMode}`}>
      <Header
        role={role}
        active={view}
        displayMode={displayMode}
        onToggleDisplayMode={() => setDisplayMode((current) => current === 'night' ? 'day' : 'night')}
        onNavigate={navigate}
        onLogout={onLogout}
      />
      <div className="content">
        {view === 'project-detail' ? (
          <ProjectDetail admin={admin} group={selectedGroup} work={selectedWork} />
        ) : view === 'project-group' ? (
          <ProjectGroup
            admin={admin}
            group={selectedGroup ?? { id: 1, name: 'Analisis de Datos y Business Intelligence', image: '', style: 'data' }}
            onOpenProject={(work) => {
              setSelectedWork(work);
              setView('project-detail');
            }}
          />
        ) : view === 'settings' && !admin ? (
          <p className="empty">Configuracion solo disponible para administrador.</p>
        ) : (
          <LandingPage
            admin={admin}
            theme={theme}
            setTheme={setTheme}
            onOpenProject={(group) => {
              setSelectedGroup(group);
              setSelectedWork(null);
              setView('project-group');
            }}
          />
        )}
      </div>
    </main>
  );
}

function App() {
  const [role, setRole] = useState<Role | null>(null);
  const [theme, setRawTheme] = useHydratedPersistentState<PortfolioTheme>('portfolio.theme', defaultTheme);
  const normalizedTheme = normalizeTheme(theme);
  const setTheme: React.Dispatch<React.SetStateAction<PortfolioTheme>> = (value) => {
    setRawTheme((current) => normalizeTheme(typeof value === 'function' ? value(normalizeTheme(current)) : value));
  };

  useEffect(() => {
    const root = document.documentElement;
    root.style.setProperty('--bg', normalizedTheme.background);
    root.style.setProperty('--panel', `rgba(${hexToRgb(normalizedTheme.panel)}, 0.72)`);
    root.style.setProperty('--panel-muted', `rgba(${hexToRgb(normalizedTheme.panelMuted)}, 0.78)`);
    root.style.setProperty('--line', normalizedTheme.line);
    root.style.setProperty('--line-strong', normalizedTheme.lineStrong);
    root.style.setProperty('--ink', normalizedTheme.ink);
    root.style.setProperty('--muted', normalizedTheme.muted);
    root.style.setProperty('--accent', normalizedTheme.accent);
    root.style.setProperty('--blue', normalizedTheme.blue);
    root.style.setProperty('--danger', normalizedTheme.danger);
    root.style.setProperty('--hover-glow', normalizedTheme.hoverGlow);
    root.style.setProperty('--hover-glow-rgb', hexToRgb(normalizedTheme.hoverGlow));
    root.style.setProperty('--glass-panel-rgb', hexToRgb(normalizedTheme.panel));
    root.style.setProperty('--glass-muted-rgb', hexToRgb(normalizedTheme.panelMuted));
    root.style.setProperty('--glass-line-rgb', hexToRgb(normalizedTheme.line));
    root.style.setProperty('--shadow-rgb', hexToRgb(normalizedTheme.blue));
    root.style.setProperty('--login-bg-image', normalizedTheme.loginBackground?.url ? `url("${normalizedTheme.loginBackground.url}")` : 'none');
  }, [normalizedTheme]);

  return role ? (
    <Dashboard role={role} theme={normalizedTheme} setTheme={setTheme} onLogout={() => setRole(null)} />
  ) : (
    <Login onEnter={setRole} />
  );
}

ReactDOM.createRoot(document.getElementById('root')!).render(<App />);
