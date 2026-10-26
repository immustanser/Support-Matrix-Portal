import * as React from 'react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { Label } from '@fluentui/react/lib/Label';
import { IPersonField } from '../models/IPersonField';
import { SupportMatrixService } from '../services/SupportMatrixService';

export interface IPersonFieldPickerProps {
  label: string;
  /** Pass a single-element array for single-person fields, or multiple for multi-person fields. */
  selected: IPersonField[];
  onChange: (people: IPersonField[]) => void;
  itemLimit?: number;
  disabled?: boolean;
  /**
   * Optional extra pool of already-known people (e.g. extracted from existing Applications
   * data) that can be searched/selected alongside the live directory, even if a person is
   * no longer returned by directory search. Selecting one of these never requires a
   * directory lookup - the provided IPersonField identity is used directly. When omitted,
   * behavior is unchanged from the directory-only picker.
   */
  localDataset?: IPersonField[];
}

const toPersona = (person: IPersonField): IPersonaProps => ({
  key: `person-${person.id}`,
  text: person.title,
  secondaryText: person.email,
  id: String(person.id)
});

const LOCAL_PERSONA_KEY_PREFIX = 'local-';

export const PersonFieldPicker: React.FC<IPersonFieldPickerProps> = ({
  label,
  selected,
  onChange,
  itemLimit,
  disabled,
  localDataset
}) => {
  const selectedPersonas = React.useMemo(() => selected.map(toPersona), [selected]);

  const onResolveSuggestions = React.useCallback(
    async (filterText: string): Promise<IPersonaProps[]> => {
      if (!filterText) {
        return [];
      }
      const suggestions = await SupportMatrixService.searchPeople(filterText);
      const directoryPersonas: IPersonaProps[] = suggestions.map((s) => ({
        key: s.key,
        text: s.text,
        secondaryText: s.secondaryText
      }));

      if (!localDataset || localDataset.length === 0) {
        return directoryPersonas;
      }

      // Avoid showing a duplicate entry for someone the live directory already returned.
      const directoryEmails = new Set(
        directoryPersonas
          .map((p) => (p.secondaryText || '').toLowerCase())
          .filter((email) => !!email)
      );

      const lowerFilter = filterText.toLowerCase();
      const seenLocalIds = new Set<number>();
      const localPersonas: IPersonaProps[] = [];
      localDataset.forEach((person) => {
        const email = (person.email || '').toLowerCase();
        if (email && directoryEmails.has(email)) {
          return;
        }
        if (seenLocalIds.has(person.id)) {
          return;
        }
        const matches =
          person.title.toLowerCase().indexOf(lowerFilter) !== -1 ||
          email.indexOf(lowerFilter) !== -1;
        if (!matches) {
          return;
        }
        seenLocalIds.add(person.id);
        localPersonas.push({
          key: `${LOCAL_PERSONA_KEY_PREFIX}${person.id}`,
          text: person.title,
          secondaryText: person.email ? `${person.email} (Existing assignment)` : 'Existing assignment'
        });
      });

      return [...directoryPersonas, ...localPersonas];
    },
    [localDataset]
  );

  const onEmptyResolveSuggestions = React.useCallback((): IPersonaProps[] => [], []);

  const onItemSelected = React.useCallback(
    async (item?: IPersonaProps): Promise<IPersonaProps> => {
      if (!item || !item.key) {
        return item as IPersonaProps;
      }
      const keyStr = String(item.key);
      if (localDataset && keyStr.indexOf(LOCAL_PERSONA_KEY_PREFIX) === 0) {
        // Sourced from existing application assignment data - the identity is already
        // known, so no directory lookup is required (and must not be required) to select it.
        const localId = Number(keyStr.substring(LOCAL_PERSONA_KEY_PREFIX.length));
        const localPerson = localDataset.filter((p) => p.id === localId)[0];
        if (localPerson) {
          const alreadySelected = selected.some((p) => p.id === localPerson.id);
          if (!alreadySelected) {
            const next = itemLimit === 1 ? [localPerson] : [...selected, localPerson];
            onChange(next);
          }
          return toPersona(localPerson);
        }
      }
      try {
        const resolved = await SupportMatrixService.resolvePerson(
          String(item.key),
          item.text || '',
          item.secondaryText || ''
        );
        const alreadySelected = selected.some((p) => p.id === resolved.id);
        if (!alreadySelected) {
          const next = itemLimit === 1 ? [resolved] : [...selected, resolved];
          onChange(next);
        }
        return toPersona(resolved);
      } catch (error) {
        console.error('PersonFieldPicker: failed to resolve person', error);
        return item;
      }
    },
    [selected, onChange, itemLimit, localDataset]
  );

  const onChangePersonas = React.useCallback(
    (items?: IPersonaProps[]): void => {
      if (!items) {
        onChange([]);
        return;
      }
      const ids = items.map((item) => (item.id ? Number(item.id) : undefined)).filter((id): id is number => !!id);
      onChange(selected.filter((p) => ids.indexOf(p.id) !== -1));
    },
    [selected, onChange]
  );

  return (
    <div>
      <Label>{label}</Label>
      <NormalPeoplePicker
        onResolveSuggestions={onResolveSuggestions}
        onEmptyResolveSuggestions={onEmptyResolveSuggestions}
        onItemSelected={onItemSelected}
        onChange={onChangePersonas}
        selectedItems={selectedPersonas}
        itemLimit={itemLimit}
        disabled={disabled}
        resolveDelay={300}
        inputProps={{ placeholder: 'Search by name or email...' }}
      />
    </div>
  );
};

export default PersonFieldPicker;
