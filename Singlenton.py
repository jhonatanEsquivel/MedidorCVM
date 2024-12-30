import tkinter as tk

class Singleton:
    _instance = None

    def __new__(cls):
        if not cls._instance:
            cls._instance = super(Singleton, cls).__new__(cls)
            cls._instance.variables = {}
            cls._instance.init_variables()
        return cls._instance

    def init_variables(self):
        # Agrega aquí todas las variables que necesites
        string_variables = ['vista_cliente', 'vista_area', 'vista_lectura',
                            'vista_emision', 'vista_informe', 'vista_dias',
                            'vista_mes1',
                            'vista_mes2',
                            'vista_mes3',
                            'vista_mes4',
                            'vista_mes5',
                            'vista_mes6',
                            'vista_nombreMes1',
                            'vista_nombreMes2',
                            'vista_nombreMes3',
                            'vista_nombreMes4',
                            'vista_nombreMes5',
                            'vista_nombreMes6', 
                            'vista_c_fijoMensual', 
                            'vista_c_energiaActivaPunta', 
                            'vista_c_energiaActivaFueraPunta',
                            'vista_c_energiaReactivaExc30', 
                            'vista_c_potenciaActivaGeneracionUsuariosPresentePunta', 
                            'vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta',
                            'vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta', 
                            'vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta', 
                            'vista_t_nombreTablero',
                            'vista_t_energiaActivaHoraFueraPuntaActual', 
                            'vista_t_energiaActivaHoraFueraPuntaAnterior', 
                            'vista_t_energiaActivaHoraPuntaActual', 
                            'vista_t_energiaActivaHoraPuntaAnterior',
                            'vista_t_evidencia1', 
                            'vista_t_maximaDemanda',
                            'vista_t_evidencia2', 
                            'vista_t_energiaReactivaInductivaActual', 
                            'vista_t_energiaReactivaInductivaAnterior',
                            'vista_t_evidencia3',
                            'vista_cantidadMedidores',
                            'vista_a_Actual', 
                            'vista_a_Anterior',
                            'vista_a_evidencia1', 
                            'vista_a_cantidadMedidores', 
                            'data_a_total',
                            'data_a_suma',
                            'data_promedio',
                            'data_sumaAB_t1',
                            'data_sumaAB_t2',
                            'data_sumaAB_total',
                            'data_sumaC_total',
                            'data_sumaD_total',
                            'data_horasPunta',
                            'data_calificacion',
                            'data_calificacionTarifaria']
        array_variables = [ 'array_t_nombreTablero',
                            'array_t_energiaActivaHoraFueraPuntaActual', 
                            'array_t_energiaActivaHoraFueraPuntaAnterior',
                            'array_energiaActivaHoraFueraPunta', 
                            'array_t_energiaActivaHoraPuntaActual', 
                            'array_t_energiaActivaHoraPuntaAnterior',
                            'array_energiaActivaHoraPunta',
                            'array_energiaActivaActual',
                            'array_energiaActivaAnterior',
                            'array_energiaActivaTotal',
                            'array_t_evidencia1', 
                            'array_t_maximaDemanda',
                            'array_t_evidencia2', 
                            'array_t_energiaReactivaInductivaActual', 
                            'array_t_energiaReactivaInductivaAnterior',
                            'array_energiaReactivaInductivaTotal',
                            'array_t_evidencia3',
                            'array_a_Actual', 
                            'array_a_Anterior',
                            'array_a_Total',
                            'array_a_evidencia1'
                              ]

        for name in string_variables:
            self.variables[name] = tk.StringVar()
            self.variables[name].set('')

        for name in array_variables:
            self.variables[name] = []  # Puedes inicializar tus listas de la manera que prefieras
        
        # Inicializar variable firma
        self.variables['firma'] = tk.BooleanVar()
        self.variables['firma'].set(False)

    def get_variable(self, variable_name):
        return self.variables.get(variable_name)

    def set_variable(self, variable_name, new_value):
        variable = self.variables.get(variable_name)
        if variable:
            if isinstance(variable, tk.StringVar):
                variable.set(new_value)
            elif isinstance(variable, list):
                # Si es una lista, limpiamos la lista y luego extendemos con los nuevos valores
                variable.clear()
                variable.extend(new_value)
            else:
                print(f"Tipo de variable no compatible para {variable_name}.")
        else:
            print(f"La variable {variable_name} no existe en el Singleton.")

    
    

