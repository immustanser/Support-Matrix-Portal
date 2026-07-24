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
}

const toPersona = (person: IPersonField): IPersonaProps => ({
  key: `person-${person.id}`,
  text: person.title,
  secondaryText: person.email,
  id: String(person.id)
});

export const PersonFieldPicker: React.FC<IPersonFieldPickerProps> = ({
  label,
  selected,
  onChange,
  itemLimit,
  disabled
}) => {
  const selectedPersonas = React.useMemo(() => selected.map(toPersona), [selected]);

  const onResolveSuggestions = React.useCallback(
    async (filterText: string): Promise<IPersonaProps[]> => {
      if (!filterText) {
        return [];
      }
      const suggestions = await SupportMatrixService.searchPeople(filterText);
      return suggestions.map((s) => ({
        key: s.key,
        text: s.text,
        secondaryText: s.secondaryText
      }));
    },
    []
  );

  const onEmptyResolveSuggestions = React.useCallback((): IPersonaProps[] => [], []);

  const onItemSelected = React.useCallback(
    async (item?: IPersonaProps): Promise<IPersonaProps> => {
      if (!item || !item.key) {
        return item as IPersonaProps;
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
    [selected, onChange, itemLimit]
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
