# Frontend Integration - Group Letter Examples

## 📋 Ejemplo Completo de Integración

### 1. Modelos TypeScript

```typescript
// types/presentation.ts

export interface SlideDetail {
  slide_number: number;
  slide_type: string;
  slide_bg_file_name: string;
  slide_description: string;
  group_name: string;        // "A|Prescreen Survivors" o "Prescreen Survivors"
  category: string;
  name: string;
  rationale: string;
  notation: string;
  kana: string;
  logo_filename: string;
  template_id: number;
  name_sub_group: string;
  group_letter: string;      // "A", "B", "C" o "" (NUEVO CAMPO)
}

export interface PresentationDetails {
  presentation_id: number;
  project: string;
  display_name: string;
  details: SlideDetail[];
}

export type GroupLetter = "A" | "B" | "C" | "";
export type VotingBehavior = "voting" | "no-voting" | "default";
export type VoteType = "positive" | "neutral" | "negative";
```

### 2. Helper Functions

```typescript
// utils/groupHelpers.ts

export interface ParsedGroupName {
  letter: string;
  name: string;
}

/**
 * Parsea el campo group_name en formato "A|Nombre"
 * @param groupName - String en formato "A|Prescreen Survivors"
 * @returns Objeto con letter y name separados
 */
export function parseGroupName(groupName: string): ParsedGroupName {
  if (!groupName) {
    return { letter: "", name: "" };
  }

  // Verificar si tiene el formato "A|Nombre"
  if (groupName.includes("|")) {
    const [letter, ...nameParts] = groupName.split("|");
    return {
      letter: letter.trim(),
      name: nameParts.join("|").trim()  // Por si hay más "|" en el nombre
    };
  }

  // Retrocompatibilidad: sin letra
  return { letter: "", name: groupName };
}

/**
 * Determina si un slide permite votación basado en la letra del grupo
 * @param groupLetter - Letra del grupo ("A", "B", "C" o "")
 * @returns true si el grupo permite votación
 */
export function canVote(groupLetter: string): boolean {
  return groupLetter === "B";
}

/**
 * Obtiene el comportamiento de votación del grupo
 * @param groupLetter - Letra del grupo
 * @returns Comportamiento: "voting", "no-voting" o "default"
 */
export function getVotingBehavior(groupLetter: string): VotingBehavior {
  switch (groupLetter) {
    case "B":
      return "voting";
    case "C":
      return "no-voting";
    case "A":
    case "":
    default:
      return "default";
  }
}

/**
 * Extrae el nombre limpio del grupo (sin la letra)
 * @param groupName - String en formato "A|Nombre" o "Nombre"
 * @returns Nombre del grupo sin letra
 */
export function getCleanGroupName(groupName: string): string {
  const { name } = parseGroupName(groupName);
  return name;
}
```

### 3. Hook Personalizado (React)

```typescript
// hooks/useSlideVoting.ts
import { useState } from 'react';
import { SlideDetail, VoteType } from '@/types/presentation';
import { canVote } from '@/utils/groupHelpers';

export function useSlideVoting(slide: SlideDetail) {
  const [currentVote, setCurrentVote] = useState<VoteType | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  // Determinar si el slide permite votación
  const votingAllowed = canVote(slide.group_letter);

  // Función para votar
  const vote = async (voteType: VoteType) => {
    if (!votingAllowed) {
      console.warn('Voting not allowed for this slide');
      return;
    }

    setIsLoading(true);

    try {
      // Aquí harías la llamada al API
      const response = await fetch(`/api/presentations/${slide.slide_number}/vote`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          slide_number: slide.slide_number,
          vote_type: voteType
        })
      });

      if (response.ok) {
        setCurrentVote(voteType);
      }
    } catch (error) {
      console.error('Error voting:', error);
    } finally {
      setIsLoading(false);
    }
  };

  return {
    votingAllowed,
    currentVote,
    isLoading,
    vote
  };
}
```

### 4. Componente de Slide (React)

```tsx
// components/SlideCard.tsx
import React from 'react';
import { SlideDetail } from '@/types/presentation';
import { useSlideVoting } from '@/hooks/useSlideVoting';
import { getCleanGroupName } from '@/utils/groupHelpers';

interface SlideCardProps {
  slide: SlideDetail;
}

export function SlideCard({ slide }: SlideCardProps) {
  const { votingAllowed, currentVote, isLoading, vote } = useSlideVoting(slide);

  return (
    <div className="slide-card">
      {/* Header con nombre del grupo */}
      <div className="slide-header">
        <span className="group-badge">
          Group {slide.group_letter || 'N/A'}
        </span>
        <h3>{getCleanGroupName(slide.group_name)}</h3>
      </div>

      {/* Contenido del slide */}
      <div className="slide-content">
        <h2>{slide.name}</h2>
        {slide.rationale && (
          <p className="rationale">{slide.rationale}</p>
        )}
        {slide.notation && (
          <p className="notation">Notation: {slide.notation}</p>
        )}
      </div>

      {/* Botones de votación (solo si está permitido) */}
      {votingAllowed && (
        <div className="voting-section">
          <h4>Cast Your Vote</h4>
          <div className="voting-buttons">
            <button
              onClick={() => vote('positive')}
              disabled={isLoading}
              className={currentVote === 'positive' ? 'active' : ''}
            >
              👍 Positive
            </button>
            <button
              onClick={() => vote('neutral')}
              disabled={isLoading}
              className={currentVote === 'neutral' ? 'active' : ''}
            >
              😐 Neutral
            </button>
            <button
              onClick={() => vote('negative')}
              disabled={isLoading}
              className={currentVote === 'negative' ? 'active' : ''}
            >
              👎 Negative
            </button>
          </div>
        </div>
      )}

      {/* Mensaje para grupos sin votación */}
      {!votingAllowed && slide.group_letter === 'C' && (
        <div className="no-voting-message">
          <p>ℹ️ Voting is not available for this group</p>
        </div>
      )}
    </div>
  );
}
```

### 5. Lista de Slides con Filtrado

```tsx
// components/SlideList.tsx
import React, { useState } from 'react';
import { SlideDetail } from '@/types/presentation';
import { SlideCard } from './SlideCard';
import { getVotingBehavior } from '@/utils/groupHelpers';

interface SlideListProps {
  slides: SlideDetail[];
}

export function SlideList({ slides }: SlideListProps) {
  const [filterByVoting, setFilterByVoting] = useState<'all' | 'voting' | 'no-voting'>('all');

  // Filtrar slides por comportamiento de votación
  const filteredSlides = slides.filter(slide => {
    if (filterByVoting === 'all') return true;

    const behavior = getVotingBehavior(slide.group_letter);

    if (filterByVoting === 'voting') {
      return behavior === 'voting';
    }

    if (filterByVoting === 'no-voting') {
      return behavior === 'no-voting' || behavior === 'default';
    }

    return true;
  });

  // Estadísticas
  const stats = {
    total: slides.length,
    voting: slides.filter(s => getVotingBehavior(s.group_letter) === 'voting').length,
    noVoting: slides.filter(s => {
      const b = getVotingBehavior(s.group_letter);
      return b === 'no-voting' || b === 'default';
    }).length
  };

  return (
    <div className="slide-list">
      {/* Filtros */}
      <div className="filters">
        <button
          onClick={() => setFilterByVoting('all')}
          className={filterByVoting === 'all' ? 'active' : ''}
        >
          All Slides ({stats.total})
        </button>
        <button
          onClick={() => setFilterByVoting('voting')}
          className={filterByVoting === 'voting' ? 'active' : ''}
        >
          Voting Enabled ({stats.voting})
        </button>
        <button
          onClick={() => setFilterByVoting('no-voting')}
          className={filterByVoting === 'no-voting' ? 'active' : ''}
        >
          No Voting ({stats.noVoting})
        </button>
      </div>

      {/* Grid de slides */}
      <div className="slides-grid">
        {filteredSlides.map(slide => (
          <SlideCard key={slide.slide_number} slide={slide} />
        ))}
      </div>
    </div>
  );
}
```

### 6. API Service (TypeScript)

```typescript
// services/presentationService.ts

export interface VoteRequest {
  slide_number: number;
  vote_type: 'positive' | 'neutral' | 'negative';
  user_id?: string;
}

export interface VoteResponse {
  success: boolean;
  slide_number: number;
  vote_type: string;
  timestamp: string;
}

export class PresentationService {
  private baseUrl = '/api/presentations';

  /**
   * Obtener detalles de una presentación
   */
  async getPresentationDetails(presentationId: number): Promise<PresentationDetails> {
    const response = await fetch(`${this.baseUrl}/${presentationId}/details`);

    if (!response.ok) {
      throw new Error('Failed to fetch presentation details');
    }

    return response.json();
  }

  /**
   * Votar en un slide
   */
  async voteOnSlide(
    presentationId: number,
    voteRequest: VoteRequest
  ): Promise<VoteResponse> {
    const response = await fetch(
      `${this.baseUrl}/${presentationId}/slides/${voteRequest.slide_number}/vote`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(voteRequest)
      }
    );

    if (!response.ok) {
      throw new Error('Failed to submit vote');
    }

    return response.json();
  }

  /**
   * Obtener votos de un slide
   */
  async getSlideVotes(presentationId: number, slideNumber: number) {
    const response = await fetch(
      `${this.baseUrl}/${presentationId}/slides/${slideNumber}/votes`
    );

    if (!response.ok) {
      throw new Error('Failed to fetch votes');
    }

    return response.json();
  }
}

export const presentationService = new PresentationService();
```

### 7. Tests Unitarios (Jest)

```typescript
// utils/groupHelpers.test.ts
import {
  parseGroupName,
  canVote,
  getVotingBehavior,
  getCleanGroupName
} from './groupHelpers';

describe('groupHelpers', () => {
  describe('parseGroupName', () => {
    it('should parse group name with letter', () => {
      expect(parseGroupName('A|Prescreen Survivors')).toEqual({
        letter: 'A',
        name: 'Prescreen Survivors'
      });
    });

    it('should handle group name without letter', () => {
      expect(parseGroupName('Prescreen Survivors')).toEqual({
        letter: '',
        name: 'Prescreen Survivors'
      });
    });

    it('should handle empty string', () => {
      expect(parseGroupName('')).toEqual({
        letter: '',
        name: ''
      });
    });

    it('should handle group name with multiple pipes', () => {
      expect(parseGroupName('B|Group|With|Pipes')).toEqual({
        letter: 'B',
        name: 'Group|With|Pipes'
      });
    });
  });

  describe('canVote', () => {
    it('should allow voting for group B', () => {
      expect(canVote('B')).toBe(true);
    });

    it('should not allow voting for group A', () => {
      expect(canVote('A')).toBe(false);
    });

    it('should not allow voting for group C', () => {
      expect(canVote('C')).toBe(false);
    });

    it('should not allow voting for empty group', () => {
      expect(canVote('')).toBe(false);
    });
  });

  describe('getVotingBehavior', () => {
    it('should return voting for group B', () => {
      expect(getVotingBehavior('B')).toBe('voting');
    });

    it('should return no-voting for group C', () => {
      expect(getVotingBehavior('C')).toBe('no-voting');
    });

    it('should return default for group A', () => {
      expect(getVotingBehavior('A')).toBe('default');
    });

    it('should return default for empty group', () => {
      expect(getVotingBehavior('')).toBe('default');
    });
  });

  describe('getCleanGroupName', () => {
    it('should extract clean name from formatted group name', () => {
      expect(getCleanGroupName('A|Prescreen Survivors')).toBe('Prescreen Survivors');
    });

    it('should return name as-is if no letter', () => {
      expect(getCleanGroupName('Prescreen Survivors')).toBe('Prescreen Survivors');
    });
  });
});
```

### 8. Estilos CSS (ejemplo)

```css
/* styles/slideCard.css */

.slide-card {
  border: 1px solid #e0e0e0;
  border-radius: 8px;
  padding: 20px;
  margin: 10px;
  background: white;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.slide-header {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 15px;
}

.group-badge {
  display: inline-block;
  padding: 4px 12px;
  border-radius: 12px;
  font-size: 12px;
  font-weight: 600;
  text-transform: uppercase;
}

/* Colores por grupo */
.group-badge[data-letter="A"] {
  background: #e3f2fd;
  color: #1976d2;
}

.group-badge[data-letter="B"] {
  background: #e8f5e9;
  color: #388e3c;
}

.group-badge[data-letter="C"] {
  background: #ffebee;
  color: #d32f2f;
}

.voting-section {
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid #e0e0e0;
}

.voting-buttons {
  display: flex;
  gap: 10px;
  margin-top: 10px;
}

.voting-buttons button {
  flex: 1;
  padding: 12px;
  border: 2px solid #e0e0e0;
  border-radius: 6px;
  background: white;
  cursor: pointer;
  transition: all 0.2s;
}

.voting-buttons button:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}

.voting-buttons button.active {
  border-color: #1976d2;
  background: #e3f2fd;
}

.voting-buttons button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.no-voting-message {
  margin-top: 15px;
  padding: 12px;
  background: #fff3cd;
  border: 1px solid #ffc107;
  border-radius: 4px;
  text-align: center;
}
```

## 🚀 Ejemplo de Integración Completa

```tsx
// pages/PresentationView.tsx
import React, { useEffect, useState } from 'react';
import { useParams } from 'react-router-dom';
import { PresentationDetails } from '@/types/presentation';
import { SlideList } from '@/components/SlideList';
import { presentationService } from '@/services/presentationService';

export function PresentationView() {
  const { presentationId } = useParams<{ presentationId: string }>();
  const [presentation, setPresentation] = useState<PresentationDetails | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    async function loadPresentation() {
      try {
        const data = await presentationService.getPresentationDetails(
          Number(presentationId)
        );
        setPresentation(data);
      } catch (err) {
        setError('Failed to load presentation');
        console.error(err);
      } finally {
        setLoading(false);
      }
    }

    loadPresentation();
  }, [presentationId]);

  if (loading) return <div>Loading...</div>;
  if (error) return <div>Error: {error}</div>;
  if (!presentation) return <div>Presentation not found</div>;

  return (
    <div className="presentation-view">
      <h1>{presentation.display_name}</h1>
      <SlideList slides={presentation.details} />
    </div>
  );
}
```

## ✅ Checklist de Integración

- [ ] Actualizar modelos TypeScript con campo `group_letter`
- [ ] Crear helper functions (`parseGroupName`, `canVote`, etc.)
- [ ] Implementar componente de slide con botones de votación
- [ ] Agregar lógica para mostrar/ocultar votación según `group_letter`
- [ ] Crear servicio API para enviar votos
- [ ] Implementar tests unitarios
- [ ] Agregar estilos CSS
- [ ] Probar con datos reales del backend
- [ ] Documentar para el equipo

---

**¡Listo para usar!** 🎉 Estos ejemplos cubren todos los casos de uso comunes.
