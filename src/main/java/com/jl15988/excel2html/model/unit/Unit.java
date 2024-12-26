package com.jl15988.excel2html.model.unit;

import com.jl15988.excel2html.constant.UnitConstant;
import com.jl15988.excel2html.enums.UnitType;
import com.jl15988.excel2html.exception.UnitException;
import com.jl15988.excel2html.utils.BeanUtil;
import lombok.Getter;
import lombok.Setter;

import java.util.Objects;

/**
 * 单位
 *
 * @author Jalon
 * @since 2024/12/13 16:40
 **/
@Getter
@Setter
public class Unit<T extends Unit<T>> {

    /**
     * 单位值
     */
    private double value;

    /**
     * 屏幕 DPI
     */
    private int dpi;

    /**
     * 单位
     */
    private final String unit;

    /**
     * 单位类
     */
    private final Class<T> clazz;

    /**
     * 单位类型
     */
    private final UnitType unitType;

    public Unit(Class<T> clazz, double value, int dpi) {
        this.clazz = clazz;
        this.value = value;
        this.dpi = dpi;

        String unit = null;
        UnitType type = null;
        if (clazz.equals(UnitPixel.class)) {
            unit = UnitConstant.PIXEL_UNIT;
            type = UnitType.PIXEL;
        } else if (clazz.equals(UnitPoint.class)) {
            unit = UnitConstant.POINT_UNIT;
            type = UnitType.POINT;
        } else if (clazz.equals(UnitMillimetre.class)) {
            unit = UnitConstant.MILLIMETRE_UNIT;
            type = UnitType.MILLIMETRE;
        } else if (clazz.equals(UnitInch.class)) {
            unit = UnitConstant.INCH_UNIT;
            type = UnitType.INCH;
        } else if (clazz.equals(UnitEmu.class)) {
            unit = UnitConstant.ENUM_UNIT;
            type = UnitType.EMU;
        } else {
            throw new UnitException("Unit class is not exist.");
        }
        this.unit = unit;
        this.unitType = type;
    }

    public Unit(Class<T> clazz, double value) {
        this(clazz, value, UnitConstant.DEFAULT_DPI);
    }

    /**
     * 转换为某个单位
     *
     * @param targetClass 目标类
     * @return 转换后的单位
     */
    public <R extends Unit<R>> Unit<R> converterTo(Class<R> targetClass) {
        String method = getConvertMethodName(targetClass);
        Double conRes = BeanUtil.exeMethod(this.clazz, method)
                .judge(m -> m.getParameterCount() == 1, this.value)
                .judge(m -> m.getParameterCount() == 2, this.value, this.dpi)
                .errorHandle(throwable -> {
                    if (Objects.nonNull(throwable)) {
                        throw new UnitException("Unit converter fail", throwable);
                    }
                })
                .to(Double.class);

        if (Objects.isNull(conRes)) {
            throw new UnitException("Unit converter fail.");
        }
        return new Unit<>(targetClass, conRes, this.dpi);
    }

    private static <R extends Unit<R>> String getConvertMethodName(Class<R> targetClass) {
        String method = null;
        if (UnitPixel.class == targetClass) {
            method = "toPixel";
        } else if (UnitPoint.class == targetClass) {
            method = "toPoint";
        } else if (UnitInch.class == targetClass) {
            method = "toInch";
        } else if (UnitMillimetre.class == targetClass) {
            method = "toMillimetre";
        } else if (UnitEmu.class == targetClass) {
            method = "toEmu";
        }
        if (method == null) {
            throw new UnitException("Unit.converterTo targetClass is invalid");
        }
        return method;
    }

    /**
     * 通过来源信息转为目标类单位
     *
     * @param resourceClass 来源类
     * @param targetClass   目标类
     * @param resourceValue 来源值
     * @param resourceDpi   来源 DPI
     * @param <R>           来源
     * @param <C>           目标
     * @return 目标单位
     */
    public static <R extends Unit<R>, C extends Unit<C>> Unit<C> converterTo(Class<R> resourceClass, double resourceValue, int resourceDpi, Class<C> targetClass) {
        Unit<R> rUnit = new Unit<>(resourceClass, resourceValue, resourceDpi);
        return rUnit.converterTo(targetClass);
    }

    /**
     * 通过来源信息转为目标类单位
     *
     * @param resourceUnit 来源单位
     * @param targetClass  目标类
     * @param <R>          来源
     * @param <C>          目标
     * @return 目标单位
     */
    public static <R extends Unit<R>, C extends Unit<C>> Unit<C> converterTo(Unit<R> resourceUnit, Class<C> targetClass) {
        return resourceUnit.converterTo(targetClass);
    }

    /**
     * 通过来源信息转为目标类
     *
     * @param resourceClass 来源类
     * @param targetClass   目标类
     * @param resourceValue 来源值
     * @param resourceDpi   来源 DPI
     * @param <R>           来源
     * @param <C>           目标
     * @return 目标单位
     */
    public static <R extends Unit<R>, C extends Unit<C>> C to(Class<R> resourceClass, double resourceValue, int resourceDpi, Class<C> targetClass) {
        Unit<C> tUnit = Unit.converterTo(resourceClass, resourceValue, resourceDpi, targetClass);
        return BeanUtil.newInstance(targetClass, new Class[]{double.class, int.class}, new Object[]{tUnit.getValue(), tUnit.getDpi()});
    }

    /**
     * 通过来源信息转为目标类
     *
     * @param resourceUnit 来源单位
     * @param targetClass  目标类
     * @param <R>          来源
     * @param <C>          目标
     * @return 目标单位
     */
    public static <R extends Unit<R>, C extends Unit<C>> C to(Unit<R> resourceUnit, Class<C> targetClass) {
        return Unit.to(resourceUnit.getClazz(), resourceUnit.getValue(), resourceUnit.getDpi(), targetClass);
    }

    public UnitPixel toPixel() {
        if (this.clazz == UnitPixel.class) return new UnitPixel(this.value, this.dpi);
        Unit<UnitPixel> pixelUnit = this.converterTo(UnitPixel.class);
        return new UnitPixel(pixelUnit.getValue(), pixelUnit.getDpi());
    }

    public UnitPoint toPoint() {
        if (this.clazz == UnitPoint.class) return new UnitPoint(this.value, this.dpi);
        Unit<UnitPoint> pointUnit = this.converterTo(UnitPoint.class);
        return new UnitPoint(pointUnit.getValue(), pointUnit.getDpi());
    }

    public UnitMillimetre toMillimetre() {
        if (this.clazz == UnitMillimetre.class) return new UnitMillimetre(this.value, this.dpi);
        Unit<UnitMillimetre> millimetreUnit = this.converterTo(UnitMillimetre.class);
        return new UnitMillimetre(millimetreUnit.getValue(), millimetreUnit.getDpi());
    }

    public UnitInch toInch() {
        if (this.clazz == UnitInch.class) return new UnitInch(this.value, this.dpi);
        Unit<UnitInch> inchUnit = this.converterTo(UnitInch.class);
        return new UnitInch(inchUnit.getValue(), inchUnit.getDpi());
    }

    public UnitEmu toEmu() {
        if (this.clazz == UnitEmu.class) return new UnitEmu(this.value, this.dpi);
        Unit<UnitEmu> emuUnit = this.converterTo(UnitEmu.class);
        return new UnitEmu(emuUnit.getValue(), emuUnit.getDpi());
    }

    public String toString() {
        return this.value + this.unit;
    }
}
