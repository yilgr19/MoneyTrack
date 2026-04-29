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
        duration: 2800,
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
      /** `base64: true` entrega píxeles al OCR sin leer `file://` (más fiable en Android/iOS que readAsStringAsync). */
      const photo = await camRef.current.takePictureAsync({
        quality: 0.85,
        skipProcessing: false,
        base64: true,
      });
      if (!photo?.uri && !(photo?.base64 && photo.base64.length > 0)) {
        onClose?.();
        return;
      }
      setAnalyze(true);
      const texto = await extraerTextoDeImagen({
        uri: photo.uri,
        base64: photo.base64,
      });
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

  /**
   * Facturas colombianas / tickets alargados: marco alto (formato papel), no cuadrado tipo QR.
   * Ocupa todo el ancho menos márgenes y casi todo el alto entre cabecera y botón inferior.
   */
  const footerBlock = Math.min(height * 0.2, 176) + Math.max(insets.bottom, spacing.sm);
  const headerBlock = insets.top + 84;
  const maxFrameH = height - headerBlock - footerBlock - spacing.sm;
  const frameW = width - spacing.sm * 2;
  const portraitMinH = frameW * 2.35;
  const portraitMaxH = frameW * 3.85;
  const frameH = Math.min(Math.max(maxFrameH, portraitMinH), portraitMaxH, maxFrameH);

  const scanTranslate = scanAnim.interpolate({
    inputRange: [0, 1],
    outputRange: [0, Math.max(frameH - 4, 24)],
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
          <View style={[styles.instrBanner, { marginTop: insets.top + 48 }]}>
            <Text style={styles.instrMain}>Documento largo · factura o ticket</Text>
            <Text style={styles.instrSub}>
              No es código QR: incluye el papel completo de arriba a abajo (aleja si hace falta). Pulsa Capturar
              para analizar todo el texto.
            </Text>
          </View>
        </View>

        <View style={styles.marcoWrap}>
          <View style={[styles.frameBox, { width: frameW, height: frameH }]}>
            <View pointerEvents="none" style={[styles.frameEdgeGlow, StyleSheet.absoluteFillObject]} />
            <View style={[styles.crThin, styles.crTLThin]} />
            <View style={[styles.crThin, styles.crTRThin]} />
            <View style={[styles.crThin, styles.crBLThin]} />
            <View style={[styles.crThin, styles.crBRThin]} />
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
              accessibilityLabel="Capturar toda la factura y analizar el texto del recibo"
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
    backgroundColor: 'rgba(0,0,0,0.12)',
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
    fontSize: 17,
    fontWeight: '800',
    color: colors.text,
    letterSpacing: -0.35,
    textAlign: 'center',
    marginBottom: spacing.xs,
    lineHeight: 22,
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
    borderRadius: radii.md,
    borderWidth: 1,
    borderColor: 'rgba(125,193,145,0.45)',
    backgroundColor: 'rgba(0,0,0,0.04)',
  },
  /** Luz suave alrededor — sensación documento sobre mesa, no “ventana QR” pequeña */
  frameEdgeGlow: {
    borderRadius: radii.md - 1,
    borderWidth: 1,
    borderColor: 'rgba(167,216,222,0.25)',
    margin: 2,
  },
  crThin: {
    position: 'absolute',
    width: 22,
    height: 22,
    borderColor: 'rgba(167,216,222,0.85)',
    zIndex: 3,
  },
  crTLThin: {
    left: 6,
    top: 6,
    borderTopWidth: 2,
    borderLeftWidth: 2,
    borderTopLeftRadius: 4,
  },
  crTRThin: {
    right: 6,
    top: 6,
    borderTopWidth: 2,
    borderRightWidth: 2,
    borderTopRightRadius: 4,
  },
  crBLThin: {
    left: 6,
    bottom: 6,
    borderBottomWidth: 2,
    borderLeftWidth: 2,
    borderBottomLeftRadius: 4,
  },
  crBRThin: {
    right: 6,
    bottom: 6,
    borderBottomWidth: 2,
    borderRightWidth: 2,
    borderBottomRightRadius: 4,
  },
  scanBeam: {
    position: 'absolute',
    left: 14,
    right: 14,
    height: 3,
    backgroundColor: 'rgba(125,209,173,0.55)',
    borderRadius: 2,
    zIndex: 2,
    ...Platform.select({
      ios: {
        shadowColor: colors.mint,
        shadowOpacity: 0.45,
        shadowRadius: 8,
      },
      android: { elevation: 2 },
    }),
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
