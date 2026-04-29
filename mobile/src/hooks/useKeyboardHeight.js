import { useEffect, useState } from 'react';
import { Keyboard, Platform } from 'react-native';

/** Altura del teclado en px (0 si está oculto). Útil para padding en ScrollView / modales. */
export function useKeyboardHeight() {
  const [height, setHeight] = useState(0);
  useEffect(() => {
    if (Platform.OS === 'web') return undefined;
    const showEvt = Platform.OS === 'ios' ? 'keyboardWillShow' : 'keyboardDidShow';
    const hideEvt = Platform.OS === 'ios' ? 'keyboardWillHide' : 'keyboardDidHide';
    const onShow = (e) => {
      const h = e?.endCoordinates?.height;
      setHeight(typeof h === 'number' && h > 0 ? h : 0);
    };
    const onHide = () => setHeight(0);
    const s = Keyboard.addListener(showEvt, onShow);
    const h = Keyboard.addListener(hideEvt, onHide);
    return () => {
      s.remove();
      h.remove();
    };
  }, []);
  return height;
}
