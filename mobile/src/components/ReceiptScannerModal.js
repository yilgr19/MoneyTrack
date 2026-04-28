import React, { useRef, useState, useCallback, useEffect } from 'react';
import {
  View,
  Text,
  StyleSheet,
  Modal,
  TouchableOpacity,
  ActivityIndicator,
  Animated,
  Easing,
  Platform,
  useWindowDimensions,
} from 'react-native';
import { useSafeAreaInsets } from 'react-native-safe-area-context';
import { CameraView, useCameraPermissions } from 'expo-camera';
import { Ionicons } from '@expo/vector-icons';
import { extraerTextoDeImagen, parseDatosTicketDesdeTexto } from '../lib/ocrRecibo';
import { colors, spacing, radii, typography } from '../theme';

/**
 * Overlay de cámara para recibo (sin nueva pantalla de navegación).
 * Tras capturar: OCR → devuelve datos sugeridos para el formulario.
 */
export default function ReceiptScannerModal({ visible, onClose, onDatosParsed }) {
  const insets = useSafeAreaInsets();
  const { width, height } = useWindowDimensions();
  const [perm, reqPerm] = useCameraPermissions();
  const camRef = useRef(null);
  const [lista, setLista] = useState(false);
  const [capturing, setCapturing] = useState(false);
  const [analyze, setAnalyze] = useState(false);

  const scanAnim = useRef(new Animated.Value(0)).current;
  const flashAnim = useRef(new Animated.Value(0)).current;

  useEffect(() => {
    if (!visible || !lista) return undefined;
    const loop = Animated.loop(
      Animated.timing(scanAnim, {
        toValue: 1,
        duration: 2200,
        easing: Easing.linear,
        useNativeDriver: true,
      })
    );
    loop.start();
    return () => {
      scanAnim.setValue(0);
      loop.stop();
    };
  }, [visible, lista, scanAnim]);

  useEffect(() => {
    if (visible) {
      reqPerm?.();
      setLista(false);
      setCapturing(false);
      setAnalyze(false);
    }
  }, [visible, reqPerm]);

  const onCamReady = useCallback(() => {
    setLista(true);
  }, []);

  const tomarYCerrar = useCallback(async () => {
    if (!camRef.current || capturing) return;
    setCapturing(true);
    try {
      const photo = await camRef.current.takePictureAsync({ quality: 0.82, skipProcessing: false });
      if (!photo?.uri) {
        onClose?.();
        return;
      }
      setAnalyze(true);
      const texto = await extraerTextoDeImagen(photo.uri);
      const datos = parseDatosTicketDesdeTexto(texto);
      const ok = datos.monto != null;
      Animated.sequence([
        Animated.timing(flashAnim, {
          toValue: ok ? 1 : 0.45,
          duration: ok ? 120 : 280,
          useNativeDriver: true,
        }),
        Animated.timing(flashAnim, {
          toValue: 0,
          duration: 320,
          useNativeDriver: true,
        }),
      ]).start();

      await new Promise((r) => setTimeout(r, ok ? 200 : 0));
      onDatosParsed?.({
        ...datos,
        textoCompleto: texto,
      });
      onClose?.();
    } catch (_) {
      onDatosParsed?.({ monto: null, establecimiento: null, fecha: null, textoCompleto: '' });
      onClose?.();
    } finally {
      setCapturing(false);
      setAnalyze(false);
      setLista(false);
    }
  }, [capturing, flashAnim, onClose, onDatosParsed]);

  const frameW = Math.min(width - spacing.xl * 2, 520);
  const frameH = frameW * 0.72;

  const scanTranslate = scanAnim.interpolate({
    inputRange: [0, 1],
    outputRange: [0, frameH - 3],
  });

  const showCam = Platform.OS !== 'web' && perm?.granted;
  const needPerm = Platform.OS !== 'web' && perm && !perm.granted && perm.canAskAgain !== false;

  if (!visible) return null;

  return (
    <Modal visible={visible} animationType="fade" transparent={false}>
      <View style={[styles.root, { paddingTop: insets.top, paddingBottom: insets.bottom }]}>
        {showCam ? (
          <CameraView
            ref={camRef}
            style={StyleSheet.absoluteFillObject}
            facing="back"
            onCameraReady={onCamReady}
            autofocus="on"
          />
        ) : (
          <View style={styles.camPlaceholder}>
            <Ionicons name="camera-outline" size={52} color={colors.textMuted} />
            <Text style={styles.camPlaceholderTit}>
              {Platform.OS === 'web'
                ? 'El escáner solo está disponible en la app.'
                : 'Permiso de cámara necesario para leer el recibo'}
            </Text>
            {needPerm ? (
              <TouchableOpacity style={styles.permBtn} onPress={() => reqPerm()}>
                <Text style={styles.permBtnTxt}>Permitir cámara</Text>
              </TouchableOpacity>
            ) : null}
          </View>
        )}

        <View style={[styles.overlayTint, analyze && styles.overlayAnalyzing]} pointerEvents="none" />

        <Animated.View pointerEvents="none" style={[styles.flashBurst, { opacity: flashAnim }]} />

        <View style={[styles.hdr, StyleSheet.absoluteFillObject, styles.hdrGrow]} pointerEvents="box-none">
          <TouchableOpacity
            style={[styles.hdrClose, { top: spacing.md + (insets.top || 0) }]}
            onPress={onClose}
            accessibilityLabel="Cerrar escáner"
          >
            <Ionicons name="close" size={30} color={colors.text} />
          </TouchableOpacity>
          <View style={[styles.instrBanner, { marginTop: insets.top + 56 }]}>
            <Text style={styles.instrMain}>Apunta al recibo</Text>
            <Text style={styles.instrSub}>
              Centra la zona del total dentro del cuadro. Luego pulsa Capturar para leer texto (OCR).
            </Text>
          </View>
        </View>

        <View style={styles.marcoWrap}>
          <View style={[styles.frameBox, { width: frameW, height: frameH }]}>
            <View style={[styles.cr, styles.crTL]} />
            <View style={[styles.cr, styles.crTR]} />
            <View style={[styles.cr, styles.crBL]} />
            <View style={[styles.cr, styles.crBR]} />
            {!analyze && (
              <Animated.View
                style={[
                  styles.scanBeam,
                  {
                    transform: [{ translateY: scanTranslate }],
                  },
                ]}
              />
            )}
          </View>
        </View>

        {!analyze ? (
          <View style={[styles.footer, { paddingBottom: Math.max(insets.bottom, spacing.lg), height: height * 0.22 }]}>
            <TouchableOpacity
              style={[styles.snapBtnOuter, capturing && { opacity: 0.65 }]}
              onPress={tomarYCerrar}
              disabled={!lista || capturing}
              accessibilityRole="button"
              accessibilityLabel="Capturar recibo y analizar texto"
            >
              <View style={styles.snapBtnInner}>
                <Ionicons name="camera" size={34} color="#0c0812" />
              </View>
              <Text style={styles.snapLabel}>Capturar</Text>
            </TouchableOpacity>
          </View>
        ) : (
          <View style={[styles.footer, styles.footerAnalyzing]}>
            <ActivityIndicator size="large" color={colors.mint} />
            <Text style={styles.analyzeTxt}>Leyendo el ticket… unos segundos</Text>
          </View>
        )}
      </View>
    </Modal>
  );
}

const styles = StyleSheet.create({
  root: {
    flex: 1,
    backgroundColor: '#000',
    justifyContent: 'space-between',
  },
  camPlaceholder: {
    ...StyleSheet.absoluteFillObject,
    alignItems: 'center',
    justifyContent: 'center',
    backgroundColor: colors.bgElevated,
    padding: spacing.xl,
  },
  camPlaceholderTit: {
    ...typography.body,
    textAlign: 'center',
    color: colors.textSecondary,
    marginTop: spacing.lg,
    lineHeight: 22,
  },
  permBtn: {
    marginTop: spacing.lg,
    paddingHorizontal: spacing.lg,
    paddingVertical: spacing.md,
    borderRadius: radii.md,
    backgroundColor: colors.accentDeep,
    borderWidth: 1,
    borderColor: colors.strokeStrong,
  },
  permBtnTxt: { fontWeight: '700', color: colors.text },
  overlayTint: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(0,0,0,0.15)',
  },
  overlayAnalyzing: { backgroundColor: 'rgba(0,0,0,0.5)' },
  flashBurst: {
    ...StyleSheet.absoluteFillObject,
    backgroundColor: 'rgba(34,197,94,0.55)',
    zIndex: 5,
  },
  hdrGrow: {
    justifyContent: 'flex-start',
    alignItems: 'stretch',
    zIndex: 6,
    pointerEvents: 'box-none',
  },
  hdrClose: {
    position: 'absolute',
    right: spacing.lg,
    zIndex: 10,
    width: 46,
    height: 46,
    borderRadius: 23,
    backgroundColor: 'rgba(12,8,24,0.65)',
    alignItems: 'center',
    justifyContent: 'center',
    borderWidth: 1,
    borderColor: 'rgba(255,255,255,0.22)',
  },
  instrBanner: {
    alignSelf: 'center',
    maxWidth: 360,
    marginHorizontal: spacing.lg,
    padding: spacing.md,
    borderRadius: radii.md,
    backgroundColor: 'rgba(18,14,26,0.88)',
    borderWidth: 1,
    borderColor: 'rgba(199,195,227,0.28)',
    zIndex: 8,
  },
  instrMain: {
    fontSize: 22,
    fontWeight: '900',
    color: colors.text,
    letterSpacing: -0.4,
    textAlign: 'center',
    marginBottom: spacing.xs,
  },
  instrSub: {
    ...typography.small,
    color: colors.textSecondary,
    lineHeight: 19,
    textAlign: 'center',
  },
  marcoWrap: {
    flex: 1,
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 4,
  },
  frameBox: {
    position: 'relative',
    overflow: 'hidden',
    borderRadius: radii.lg,
  },
  cr: {
    position: 'absolute',
    width: 32,
    height: 32,
    borderColor: '#7DC191',
  },
  crTL: { left: -1, top: -1, borderTopWidth: 4, borderLeftWidth: 4, borderRadius: radii.sm },
  crTR: { right: -1, top: -1, borderTopWidth: 4, borderRightWidth: 4, borderRadius: radii.sm },
  crBL: { left: -1, bottom: -1, borderBottomWidth: 4, borderLeftWidth: 4, borderRadius: radii.sm },
  crBR: { right: -1, bottom: -1, borderBottomWidth: 4, borderRightWidth: 4, borderRadius: radii.sm },
  scanBeam: {
    position: 'absolute',
    left: 10,
    right: 10,
    height: 2,
    backgroundColor: 'rgba(167,216,222,0.95)',
    shadowColor: colors.mint,
    shadowOpacity: 0.95,
    shadowRadius: 6,
    shadowOffset: { width: 0, height: 0 },
    zIndex: 2,
  },
  footer: {
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: 9,
    backgroundColor: 'rgba(12,8,24,0.72)',
    borderTopWidth: 1,
    borderTopColor: 'rgba(199,195,227,0.12)',
    minHeight: 120,
  },
  footerAnalyzing: {
    gap: spacing.md,
    flexDirection: 'column',
  },
  snapBtnOuter: { alignItems: 'center', justifyContent: 'center' },
  snapBtnInner: {
    width: 78,
    height: 78,
    borderRadius: 39,
    backgroundColor: colors.mint,
    alignItems: 'center',
    justifyContent: 'center',
    marginBottom: spacing.sm,
    borderWidth: 4,
    borderColor: 'rgba(255,255,255,0.92)',
    ...Platform.select({
      ios: {
        shadowColor: '#fff',
        shadowOpacity: 0.35,
        shadowRadius: 12,
      },
      android: { elevation: 10 },
    }),
  },
  snapLabel: { color: colors.text, fontWeight: '800', fontSize: 15, letterSpacing: 0.3 },
  analyzeTxt: { ...typography.body, marginTop: spacing.sm, fontWeight: '600', color: colors.accentBright },
});
